from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shlex
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime
from pathlib import Path
from typing import Iterable


WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"


def site_packages_compatible(site_packages: Path) -> bool:
    """Avoid loading compiled wheels from a venv made for another Python ABI."""
    if not site_packages.exists():
        return False
    pydantic_core = site_packages / "pydantic_core"
    if not pydantic_core.exists():
        return True
    abi_tag = f"cp{sys.version_info.major}{sys.version_info.minor}"
    compiled = list(pydantic_core.glob("_pydantic_core*.pyd"))
    return not compiled or any(abi_tag in path.name for path in compiled)


# Fallback for when this script is invoked with an interpreter other than the
# project .venv's own python.exe (which already has these on sys.path).
if site_packages_compatible(VENV_SITE_PACKAGES):
    sys.path.insert(0, str(VENV_SITE_PACKAGES))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from agent_config import apply_secret_defaults, load_agent_env

try:
    from anthropic import Anthropic
except ModuleNotFoundError:
    class Anthropic:  # type: ignore[no-redef]
        def __init__(self, *_args, **_kwargs):
            raise RuntimeError("anthropic package is not installed. Install it with: pip install -r requirements.txt")

import llm_client


# 日誌の元メモを置く共有rawtextフォルダ（旧OpenAI_Agent側）。orchestrator の DEFAULT_RAWTEXT_DIR と同一。
# agent.ps1/orchestrator 経由ではこのパスが引数で明示的に渡されるが、summarize_note5.py を
# 直接パス指定なしで実行した場合もこの共有フォルダを既定にする（ローカル rawtext\ ではなく）。
# RAWTEXT_DIR 環境変数で上書き可能。
DEFAULT_SHARED_RAWTEXT = (
    r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)"
    r"\0000000000OpenAI_Agent\OpenAI-Agent\rawtext"
)
DEFAULT_INPUT = os.getenv("RAWTEXT_DIR", DEFAULT_SHARED_RAWTEXT)
DEFAULT_OUTPUT_DIR = "日誌"
DEFAULT_STATE_DIR = ".note_agent_state"
DEFAULT_MODEL = "claude-sonnet-5"
DEFAULT_ONENOTE_NOTEBOOK_NAME = "FarEasternTribe"
DEFAULT_ONENOTE_SECTION_NAME = "日誌"
GOOGLE_TASK_BATCH_SIZE = 5
GOOGLE_TASK_BATCH_PAUSE_SECONDS = 1.0


# 研究OS コマンド仕様 Version 1.0
# すべての命令は @ または ＠ で始めます。
# 現時点で実処理するもの: @todo, @python, @ask, @命令
# それ以外は専用キュー/Markdownに保存し、後続スクリプトで処理できる形にします。
SUPPORTED_AT_COMMANDS = {
    "todo": "ToDo追加",
    "python": "Python実行",
    "ask": "ChatGPT解析",
    "命令": "命令実行",
    "paper": "論文検索",
    "news": "ニュース検索",
    "mail": "メール作成",
    "onenote": "OneNote同期",
    "experiment": "実験ログ",
    "idea": "アイデア保存",
    "memo": "重要メモ",
    "reference": "文献登録",
    "schedule": "スケジュール登録",
}

QUEUE_FILE_BY_COMMAND = {
    "ask": "ask_queue.jsonl",
    "paper": "paper_search_queue.jsonl",
    "news": "news_search_queue.jsonl",
    "mail": "mail_draft_queue.jsonl",
    "onenote": "onenote_sync_queue.jsonl",
    "experiment": "experiment_log_queue.jsonl",
    "idea": "idea_log_queue.jsonl",
    "memo": "important_memo_queue.jsonl",
    "reference": "reference_queue.jsonl",
    "schedule": "schedule_queue.jsonl",
}

MARKDOWN_LOG_BY_COMMAND = {
    "ask": "ask_requests.md",
    "paper": "paper_requests.md",
    "news": "news_requests.md",
    "mail": "mail_drafts.md",
    "onenote": "onenote_sync.md",
    "experiment": "experiment_log.md",
    "idea": "ideas.md",
    "memo": "important_memos.md",
    "reference": "references.md",
    "schedule": "schedule_requests.md",
}

INSTRUCTION_COMMAND_ALIASES = {
    "タスク": "todo",
    "todo": "todo",
    "python": "python",
    "ask": "ask",
    "質問": "ask",
    "解析": "ask",
    "命令": "命令",
    "paper": "paper",
    "news": "news",
    "mail": "mail",
    "onenote": "onenote",
    "experiment": "experiment",
    "idea": "idea",
    "memo": "memo",
    "reference": "reference",
    "schedule": "schedule",
}

INSTRUCTION_COMMAND_PATTERN = re.compile(
    r"^\s*("
    + "|".join(re.escape(name) for name in sorted(INSTRUCTION_COMMAND_ALIASES, key=len, reverse=True))
    + r")(?:\s+(.*))?$",
    re.IGNORECASE,
)

NATURAL_ONENOTE_KEYWORDS = ("onenote", "one note", "OneNote", "Onenote", "ノート", "日誌セクション")
NATURAL_ONENOTE_ACTIONS = ("追加", "転記", "貼り付け", "貼付", "保存", "新しいノート", "新規ノート")


COMMAND_BLOCK_PATTERNS = [
    re.compile(r"\[COMMAND\](.*?)\[/COMMAND\]", re.IGNORECASE | re.DOTALL),
    re.compile(r"\[命令\](.*?)\[/命令\]", re.DOTALL),
    re.compile(r"^T\s*$\n(.*?)\n^T\s*$", re.MULTILINE | re.DOTALL),
]

TODO_LINE_PATTERN = re.compile(
    r"^\s*(?:[-*・□☐✅]?\s*)?(?:TODO|ToDo|To do|To d 0|タスク)[:：]?\s*(.*)$|^\s*(?:[-*・□☐✅]?\s*)?やること[:：]\s*(.*)$",
    re.IGNORECASE,
)


def is_todo_continuation_after_blank(line: str) -> bool:
    stripped = line.strip()
    return bool(stripped) and bool(re.match(r"^(?:[-*・□☐✅])", stripped))


def is_todo_end_marker(line: str) -> bool:
    stripped = line.strip().casefold()
    return "ここまで" in stripped and "todo" in stripped


class DebouncedEventHandler:
    """watchdog がある場合だけ使う、保存直後の連続イベントをまとめるハンドラ。"""

    def __init__(self, watched_paths: set[Path], callback, debounce_seconds: float):
        from watchdog.events import FileSystemEventHandler

        class Handler(FileSystemEventHandler):
            def __init__(self, outer):
                self.outer = outer
                self.last_called = 0.0

            def on_modified(self, event):
                self.outer.handle_event(event, self)

            def on_created(self, event):
                self.outer.handle_event(event, self)

        self._handler_cls = Handler
        self.watched_paths = {p.resolve() for p in watched_paths}
        self.callback = callback
        self.debounce_seconds = debounce_seconds

    def make_handler(self):
        return self._handler_cls(self)

    def handle_event(self, event, handler) -> None:
        if event.is_directory:
            return

        event_path = Path(event.src_path).resolve()
        if event_path not in self.watched_paths:
            return

        now = time.time()
        if now - handler.last_called < self.debounce_seconds:
            return

        handler.last_called = now
        time.sleep(self.debounce_seconds)
        self.callback(reason=f"file updated: {event_path.name}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="テキストメモを日誌Markdown化し、研究OS v1.0 の @ コマンドを処理します。"
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=DEFAULT_INPUT,
        help=f"読み込むテキストメモ、またはテキストメモ入りフォルダ。省略時は {DEFAULT_INPUT}",
    )
    parser.add_argument(
        "--date",
        default=date.today().isoformat(),
        help="日誌の日付。例: 2026-06-30。省略時は今日",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"出力先フォルダ。省略時は {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="同名の日誌ファイルがある場合に上書きする",
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"Claude APIで使うモデル。省略時は {DEFAULT_MODEL}",
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="入力ファイルの更新を監視し、更新があったらその都度処理する",
    )
    parser.add_argument(
        "--poll",
        action="store_true",
        help="watchdogを使わず、一定間隔で更新確認する",
    )
    parser.add_argument(
        "--interval-minutes",
        type=float,
        default=30.0,
        help="--poll時の確認間隔。省略時は30分",
    )
    parser.add_argument(
        "--debounce-seconds",
        type=float,
        default=3.0,
        help="ファイル保存直後に処理開始するまでの待機秒数。省略時は3秒",
    )
    parser.add_argument(
        "--delta-only",
        action="store_true",
        help="前回処理後に増えた差分だけを要約する。追記型メモ向け",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="差分状態を無視して、入力全体を必ず処理する。更新なし判定が疑わしい時に使う",
    )
    parser.add_argument(
        "--state-dir",
        default=DEFAULT_STATE_DIR,
        help=f"差分処理や更新検知の状態保存フォルダ。省略時は {DEFAULT_STATE_DIR}",
    )
    parser.add_argument(
        "--append-output",
        action="store_true",
        help="既存の日誌ファイルに追記する。監視＋差分処理で推奨",
    )
    parser.add_argument(
        "--extract-commands",
        action="store_true",
        help="[COMMAND]...[/COMMAND]、[命令]...[/命令]、T...T の命令候補を抽出して保存する",
    )
    parser.add_argument(
        "--extract-todos",
        action="store_true",
        help="ToDo候補を抽出して保存する",
    )
    parser.add_argument(
        "--execute-commands",
        action="store_true",
        help="@todo、@python、@ask、@命令/＠命令 を実行する。@pythonは作業フォルダ内の.pyのみ実行",
    )
    parser.add_argument(
        "--commands-only",
        action="store_true",
        help="日誌要約を作らず、入力内の@コマンドだけ実行する",
    )
    parser.add_argument(
        "--local-summary",
        dest="local_summary",
        action="store_true",
        default=True,
        help="Claude APIを使わず、ローカル抽出だけで日誌Markdownを作る（既定）",
    )
    parser.add_argument(
        "--api-summary",
        dest="local_summary",
        action="store_false",
        help="明示的にClaude APIを使って日誌Markdownを作る（課金あり・オプトイン）",
    )
    parser.add_argument(
        "--skip-ask",
        action="store_true",
        help="@askを外部APIへ送らず、キュー/ログに保存してスキップする",
    )
    parser.add_argument(
        "--sync-google-todos",
        action="store_true",
        default=os.getenv("GOOGLE_TODO_SYNC", "").lower() in {"1", "true", "yes", "on"},
        help="@todoをGoogle ToDo/Google Tasksへ即時登録する。失敗時はキューに保存",
    )
    parser.add_argument(
        "--no-sync-google-todos",
        dest="sync_google_todos",
        action="store_false",
        help="GOOGLE_TODO_SYNCが有効でも、Google Tasksへ同期しない",
    )
    parser.add_argument(
        "--manual-google-tasks",
        action="store_true",
        help="@todoはGoogleへ送らず、手動投入用TodoListだけを作る",
    )
    parser.add_argument(
        "--google-credentials",
        default=os.getenv("GOOGLE_TASKS_CREDENTIALS", "credentials.json"),
        help="Google OAuthクライアントJSON。省略時は GOOGLE_TASKS_CREDENTIALS または credentials.json",
    )
    parser.add_argument(
        "--google-token",
        default=os.getenv("GOOGLE_TASKS_TOKEN", "token_google_tasks.json"),
        help="Google Tasks OAuthトークン保存先。省略時は GOOGLE_TASKS_TOKEN または token_google_tasks.json",
    )
    parser.add_argument(
        "--google-tasklist",
        default=os.getenv("GOOGLE_TASKS_LIST_ID", "@default"),
        help="追加先Google TasksリストID。省略時は @default",
    )
    parser.add_argument(
        "--onenote-section",
        default=os.getenv("ONENOTE_SECTION_NAME", DEFAULT_ONENOTE_SECTION_NAME),
        help=f"@onenoteの追加先OneNoteセクション名。省略時は {DEFAULT_ONENOTE_SECTION_NAME}",
    )
    parser.add_argument(
        "--python-timeout",
        type=int,
        default=120,
        help="@pythonで実行するスクリプトのタイムアウト秒数。省略時は120秒",
    )
    parser.add_argument(
        "--show-rules",
        action="store_true",
        help="研究OS @コマンド仕様を表示して終了する",
    )
    parser.add_argument(
        "--no-initial-run",
        action="store_true",
        help="--watch開始時に即処理せず、次の更新まで待つ",
    )
    return parser.parse_args()


def local_extractive_summary(note: str, journal_date: str) -> str:
    lines = [line.strip() for line in note.replace("\r\n", "\n").splitlines() if line.strip()]

    def pick(keywords: tuple[str, ...], limit: int = 12) -> list[str]:
        picked = []
        for line in lines:
            if any(keyword.casefold() in line.casefold() for keyword in keywords):
                picked.append(line)
            if len(picked) >= limit:
                break
        return picked

    def unique(items: Iterable[str], limit: int) -> list[str]:
        picked = []
        seen = set()
        for item in items:
            normalized = re.sub(r"\s+", " ", item).strip().casefold()
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            picked.append(item)
            if len(picked) >= limit:
                break
        return picked

    def pick_experiment_conditions(limit: int = 16) -> list[str]:
        condition_keywords = (
            "mg", "g", "ml", "mL", "mmol", "mol/L", "M ", "℃", "°C", "oc", "oC",
            "reaction", "start", "reflux", "overnight", "rt", "room temp",
            "投入", "入れた", "加熱", "攪拌", "撹拌", "還流", "処理", "nmr", "NMR",
            "溶媒", "アセニト", "アセトニトリル", "スルトン",
        )
        picked: list[str] = []
        for index, line in enumerate(lines):
            lowered = line.casefold()
            has_condition = any(keyword.casefold() in lowered for keyword in condition_keywords)
            has_time = re.search(r"\b\d{1,2}[:：]\d{2}\b", line) is not None
            if not has_condition and not has_time:
                continue
            if index > 0 and ("reaction" in lowered or "start" in lowered):
                picked.append(lines[index - 1])
            picked.append(line)
        return unique(picked, limit)

    research = unique(
        pick(("研究", "実験", "合成", "STM", "測定", "試薬", "反応", "PPT", "論文"), limit=16)
        + pick_experiment_conditions(limit=16),
        limit=24,
    )
    todos = extract_todo_candidates(note)
    if not todos:
        todos = pick(("やる", "todo", "TODO", "確認", "作成", "掃除", "継続", "始める"), limit=10)
    mood = pick(("気分", "疲", "達成感", "ストレス", "帰", "睡眠", "体調"), limit=8)
    friction = pick(("失敗", "見つからない", "未完成", "ゴチャゴチャ", "早めに帰", "延期", "分からない"), limit=8)
    positive = pick(("楽しみ", "おもしろい", "育って", "平和", "できそう", "継続"), limit=8)
    next_actions = pick(("次", "明日", "今日中", "始め", "作成", "確認", "整理", "掃除", "発注"), limit=10)

    def bullet(items: list[str], empty: str = "記載なし") -> str:
        if not items:
            return f"- {empty}"
        return "\n".join(f"- {item}" for item in items)

    self_analysis = [
        "記録を継続し、仕組み化しようとしている傾向があります。" if pick(("記録", "ローテーション", "エージェント"), 3) else "",
        "研究・実験の停滞感と、次の作業へ進めたい気持ちが同時に出ています。" if friction and research else "",
        "タスク数が多いため、実験・事務・記録を分けて優先順位を付けると動きやすそうです。" if len(todos) >= 5 else "",
        "手書きとデジタル記録の役割分担を見直している日です。" if pick(("手書き", "Onenote", "実験ノート", "PPT"), 5) else "",
    ]
    self_analysis = [item for item in self_analysis if item]

    detailed_analysis = [
        "仮説: 研究作業そのものに加えて、記録・整理・Agent運用の整備が並行して走っており、認知負荷が高くなっています。",
        "リスク: TODOが多い日に実験を増やすと、記録漏れ、発注漏れ、後処理の先送りが起きやすくなります。" if len(todos) >= 5 else "",
        "次の一手: まずGoogle Tasksに入ったTODOを、今日必須・今日できれば・明日以降に分けるとよさそうです。",
        "次の一手: 実験ノート、OneNote、PPTは完全統一よりも、紙=その場の一次記録、OneNote=検索可能な保管、PPT=報告用の整理と役割を分けると運用しやすそうです。" if pick(("実験ノート", "Onenote", "PPT"), 5) else "",
    ]
    detailed_analysis = [item for item in detailed_analysis if item]

    overview = " ".join(lines[:5]) if lines else "入力メモが空です。"
    return "\n".join(
        [
            f"# {journal_date} 生活・研究ログ整理",
            "",
            "## 1. ローカル要約",
            overview,
            "",
            "## 2. 研究・実験メモ",
            bullet(research),
            "",
            "## 3. TODO候補",
            bullet(todos),
            "",
            "## 4. 身体・気持ち・生活",
            bullet(mood),
            "",
            "## 5. 確認事項",
            "- この日誌は `--local-summary` により外部APIを使わずローカル抽出で作成しました。",
            "- 詳細な解釈や深い分析が必要な場合は、外部API送信可否を別途確認してください。",
            "",
            "## 6. ポジティブ要素・進んでいること",
            bullet(positive),
            "",
            "## 7. 自己分析",
            bullet(self_analysis, "ローカル抽出では明確な自己分析材料を検出できませんでした。"),
            "",
            "## 8. 停滞・リスク要因",
            bullet(friction),
            "",
            "## 9. 次の行動案",
            bullet(next_actions),
            "",
            "## 10. ChatGPTによる詳細分析（ローカル代替）",
            bullet(detailed_analysis),
        ]
    )


def build_prompt(note: str, journal_date: str, delta_only: bool = False) -> str:
    target = "追記された新規メモ" if delta_only else "一日のテキストメモ"
    return f"""
あなたは、研究者の日誌・研究メモ・生活ログを整理する専属AI秘書です。
以下の{target}を、後から読み返しやすい日本語の日誌Markdownに整理してください。

重要な方針:
- 原文に誤字・脱字・音声入力由来の変換ミスがある場合は、自然な日本語に修正してください。
- ただし、意味が不明な部分は勝手に補わず、「意味不明瞭」「確認したいこと」として分けてください。
- 事実、推測、ChatGPTによる分析を明確に分けてください。
- メモに書かれていないことを事実として断定しないでください。
- 研究、生活、体調、食事、対人、気持ち、翌日の確認事項を読みやすく整理してください。
- 箇条書き中心でよいですが、日誌として自然に読めるようにしてください。
- 最後の「ChatGPTによる詳細分析」は長くても構いません。できるだけ詳しく、深く分析してください。
- 分析では、研究の進め方、心理状態、生活リズム、対人ストレス、翌日の行動設計まで含めてください。
- 必要なら「仮説」「リスク」「次の一手」に分けてください。
- 命令ブロックやToDoがある場合は、実行済みと誤解されないよう「候補」として整理してください。

出力形式:

# {journal_date} 生活・研究ログ整理

## 1. 誤字修正後の要約
原文の意味を保ちながら、誤字・脱字・表記ゆれを修正して、短く自然な文章でまとめてください。

## 2. 研究
- 今日行った研究・実験・解析
- 得られた結果
- 未完了のこと
- 次に確認すべきこと

## 3. 身体の状態
- 睡眠
- 疲労
- 体調
- 気分への影響

## 4. 食事
- 食べたもの
- 食欲
- 体調との関係

## 5. 生活・移動
- 移動
- 作業環境
- 時間の使い方

## 6. 対人・気持ちの整理
- 対人関係で起きたこと
- 感情の動き
- ストレス要因
- 事実と解釈の分離

## 7. 自己分析
- 今日の自分の行動パターン
- 集中できた点
- 乱れた点
- 繰り返し出ている傾向

## 8. TODO候補
- 研究
- 生活
- 連絡・事務
- 体調管理

## 9. 命令候補・確認したいこと
意味が曖昧な記述、追加で確認すべき点、命令候補を列挙してください。

## 10. ChatGPTによる詳細分析
以下を含めて、できるだけ詳しく分析してください。

### 10.1 今日の全体像
### 10.2 研究面の分析
### 10.3 生活・体調面の分析
### 10.4 心理面の分析
### 10.5 対人関係・ストレス構造の分析
### 10.6 明日に向けた行動戦略
### 10.7 注意すべきリスク
### 10.8 良かった点
### 10.9 次の一手

{target}:
{note}
"""


def resolve_path(path_text: str) -> Path:
    path = Path(path_text)
    if not path.is_absolute():
        path = WORKSPACE_DIR / path
    return path


def iter_text_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]
    if not input_path.is_dir():
        raise FileNotFoundError(f"入力ファイルまたはフォルダが見つかりません: {input_path}")
    text_files = sorted(input_path.glob("*.txt"))
    if not text_files:
        raise FileNotFoundError(f"フォルダ内に .txt ファイルがありません: {input_path}")
    return text_files


def read_note_text(input_path: Path) -> str:
    text_files = iter_text_files(input_path)
    chunks = []
    for text_file in text_files:
        content = text_file.read_text(encoding="utf-8-sig").strip()
        if not content:
            continue
        if len(text_files) == 1:
            chunks.append(content)
        else:
            chunks.append(f"--- {text_file.name} ---\n{content}")
    if not chunks:
        raise ValueError(f"読み込み対象の .txt ファイルがすべて空です: {input_path}")
    return "\n\n".join(chunks)


def build_raw_text_appendix(input_path: Path, note_text: str) -> str:
    text_files = iter_text_files(input_path)
    label = input_path.name if input_path.is_file() else str(input_path)
    longest_backtick_run = max((len(run) for run in re.findall(r"`+", note_text)), default=0)
    fence = "`" * max(3, longest_backtick_run + 1)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_list = "\n".join(
        "  - "
        f"{text_file.name} "
        f"(updated: {datetime.fromtimestamp(text_file.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')})"
        for text_file in text_files
    )
    return (
        "\n\n---\n\n"
        "## 生データ\n\n"
        f"- Source: `{label}`\n"
        f"- Journal generated at: {generated_at}\n"
        f"- Files:\n{file_list}\n\n"
        f"{fence}text\n"
        f"{note_text.strip()}\n"
        f"{fence}\n"
    )


def state_key(input_path: Path) -> str:
    return hashlib.sha256(str(input_path.resolve()).encode("utf-8")).hexdigest()[:16]


def state_file_path(state_dir: Path, input_path: Path) -> Path:
    return state_dir / f"{state_key(input_path)}.json"


def load_state(state_dir: Path, input_path: Path) -> dict:
    state_file = state_file_path(state_dir, input_path)
    if not state_file.exists():
        return {}
    try:
        return json.loads(state_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_state(state_dir: Path, input_path: Path, state: dict) -> None:
    state_dir.mkdir(exist_ok=True)
    state_file_path(state_dir, input_path).write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def get_delta_text(note: str, state: dict) -> tuple[str, dict]:
    previous_length = int(state.get("processed_length", 0) or 0)
    previous_hash = state.get("full_hash")
    previous_prefix_hash = state.get("processed_prefix_hash")
    current_hash = hashlib.sha256(note.encode("utf-8")).hexdigest()

    # 追記型メモなら、前回長さ以降だけを処理する。
    is_append_only = False
    if previous_length and len(note) >= previous_length:
        if previous_prefix_hash:
            current_prefix_hash = hashlib.sha256(note[:previous_length].encode("utf-8")).hexdigest()
            is_append_only = current_prefix_hash == previous_prefix_hash
        else:
            # 古い状態ファイルには prefix hash がないため、初回だけ従来どおり追記扱いにする。
            is_append_only = True

    if is_append_only:
        delta = note[previous_length:].strip()
    else:
        # ファイルが短くなった、途中編集された、または初回の場合は全体を処理。
        delta = note.strip()

    new_state = {
        **state,
        "processed_length": len(note),
        "full_hash": current_hash,
        "processed_prefix_hash": current_hash,
        "previous_hash": previous_hash,
        "last_processed_at": datetime.now().isoformat(timespec="seconds"),
    }
    return delta, new_state


def get_delta_command_text(note: str, state: dict, delta_text: str) -> str:
    """差分が既存の @ コマンドブロック内から始まる場合、直前の @ 行から命令だけ読み直す。"""
    previous_length = int(state.get("processed_length", 0) or 0)
    if not previous_length or previous_length >= len(note):
        return delta_text

    prefix = note[:previous_length]
    command_start = None
    for match in AT_COMMAND_PATTERN.finditer(prefix):
        command_start = match.start()

    if command_start is None:
        return delta_text

    command_text = note[command_start:]
    next_command = AT_COMMAND_PATTERN.search(command_text, pos=1)
    if next_command and next_command.start() < previous_length - command_start:
        return delta_text

    return command_text


def extract_command_blocks(text: str) -> list[str]:
    commands = []
    for pattern in COMMAND_BLOCK_PATTERNS:
        for match in pattern.finditer(text):
            command = match.group(1).strip()
            if command:
                commands.append(command)
    return commands


def extract_todo_candidates(text: str) -> list[str]:
    todos = []
    lines = text.splitlines()
    in_todo_section = False
    saw_blank_in_todo_section = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            if in_todo_section:
                saw_blank_in_todo_section = True
            continue
        if in_todo_section and is_todo_end_marker(stripped):
            in_todo_section = False
            saw_blank_in_todo_section = False
            continue
        if is_command_boundary(stripped):
            in_todo_section = False
            saw_blank_in_todo_section = False
            continue

        at_match = AT_COMMAND_PATTERN.match(stripped)
        if at_match:
            name = at_match.group(1).casefold().strip()
            first_arg = (at_match.group(2) or "").strip()
            in_todo_section = name == "todo"
            saw_blank_in_todo_section = False
            if in_todo_section and first_arg:
                todos.append(first_arg)
            continue

        match = TODO_LINE_PATTERN.match(stripped)
        if match:
            in_todo_section = True
            saw_blank_in_todo_section = False
            rest = (match.group(1) or match.group(2) or "").strip()
            if rest:
                todos.append(rest)
            continue

        if in_todo_section:
            if saw_blank_in_todo_section and not is_todo_continuation_after_blank(stripped):
                in_todo_section = False
                saw_blank_in_todo_section = False
                continue
            saw_blank_in_todo_section = False
            if stripped.startswith(("-", "*", "・", "□", "☐")):
                todos.append(stripped.lstrip("-*・□☐ ").strip())
            elif re.fullmatch(r"[-ーｰ－―\s]+", stripped):
                in_todo_section = False
            elif re.match(r"^[A-ZＡ-Ｚa-zａ-ｚ].{0,20}:$", stripped):
                in_todo_section = False
            elif 1 <= len(stripped) <= 120:
                todos.append(stripped)

    # 「欠席学生にレポート課題送る」のような単独行も拾う。
    # ただし @ask など別コマンドの本文は TODO 候補に混ぜない。
    in_non_todo_command = False
    for line in lines:
        stripped = line.strip(" ・-*□☐\t")
        raw_stripped = line.strip()
        if not raw_stripped:
            continue
        if is_command_boundary(raw_stripped):
            in_non_todo_command = False
            continue
        at_match = AT_COMMAND_PATTERN.match(raw_stripped)
        if at_match:
            in_non_todo_command = at_match.group(1).casefold().strip() != "todo"
            continue
        if in_non_todo_command:
            continue
        if any(keyword in stripped for keyword in ("する", "送る", "記入", "修正", "実行", "チェック", "測定", "出席")):
            if 3 <= len(stripped) <= 80 and stripped not in todos:
                todos.append(stripped)

    return todos



AT_COMMAND_PATTERN = re.compile(r"^\s*[@＠]+([^\s@＠]+)(?:\s+(.*))?$", re.IGNORECASE | re.MULTILINE)
COMMAND_BOUNDARY_PATTERN = re.compile(r"^\s*---(?:\s+.+\s+---)?\s*$")


def is_command_boundary(line: str) -> bool:
    return bool(COMMAND_BOUNDARY_PATTERN.match(line.strip()))


def parse_at_commands(text: str) -> list[dict]:
    """@todo / @python / @ask / @命令 をメモ本文から抽出する。

    書き方:
    - @todo STM測定継続
    - @todo\nSTM測定継続\n製品チェック
    - @python daily_paper_search.py --keyword CO2RR
    - @ask\n360Hzノイズについて考察して
    - ＠命令\npython daily_paper_search.py
    """
    commands: list[dict] = []
    # Some voice/OCR inputs turn the full-width ＠Todo marker into "& Todo".
    # Preserve that established journal input as a Todo marker as well.
    lines = [
        re.sub(r"^\s*[&＆]\s*todo\b", "@todo", line, flags=re.IGNORECASE)
        for line in text.splitlines()
    ]
    i = 0
    while i < len(lines):
        line = lines[i]
        match = AT_COMMAND_PATTERN.match(line)
        if not match:
            i += 1
            continue

        name = match.group(1).lower().strip()
        first_arg = (match.group(2) or "").strip()
        i += 1

        block_lines = []
        saw_blank = False
        while i < len(lines):
            next_line = lines[i]
            if AT_COMMAND_PATTERN.match(next_line) or is_command_boundary(next_line):
                break
            if name == "todo":
                if is_todo_end_marker(next_line):
                    break
                if not next_line.strip():
                    saw_blank = True
                    block_lines.append(next_line)
                    i += 1
                    continue
                if saw_blank and not is_todo_continuation_after_blank(next_line):
                    break
                saw_blank = False
            block_lines.append(next_line)
            i += 1

        block = "\n".join(block_lines).strip()
        body = first_arg if first_arg else block
        if body:
            commands.append({"name": name, "body": body})

    return commands


def looks_like_onenote_instruction(line: str) -> bool:
    lowered = line.lower()
    has_onenote_target = any(keyword.lower() in lowered for keyword in NATURAL_ONENOTE_KEYWORDS)
    has_action = any(action in line for action in NATURAL_ONENOTE_ACTIONS)
    mentions_markdown = ".md" in lowered or "markdown" in lowered or "マークダウン" in line
    return has_onenote_target and has_action and (mentions_markdown or "日誌" in line or "ノート" in line)


def normalize_instruction_body(body: str) -> str:
    """@命令 / ＠命令 の本文を既存の @ コマンド形式へ寄せる。

    例:
    ＠命令
    todo STM測定継続
    ask 360Hzノイズについて考察
    """
    normalized_lines = []
    for line in body.splitlines():
        if AT_COMMAND_PATTERN.match(line):
            normalized_lines.append(line)
            continue

        match = INSTRUCTION_COMMAND_PATTERN.match(line)
        if not match:
            if looks_like_onenote_instruction(line):
                normalized_lines.append(f"@onenote {line.strip()}")
                continue
            normalized_lines.append(line)
            continue

        raw_name = match.group(1).lower().strip()
        name = INSTRUCTION_COMMAND_ALIASES.get(raw_name, raw_name)
        rest = (match.group(2) or "").strip()
        normalized_lines.append(f"@{name} {rest}".rstrip())

    return "\n".join(normalized_lines).strip()


def command_hash(command: dict, current_output_file: Path | None = None) -> str:
    name = command.get("name", "")
    body = command.get("body", "")
    output_part = ""
    # @onenote は同じ命令文でも、生成されたMarkdownごとに実行する。
    # @命令 は自然文から @onenote に展開されることがあるため、出力ファイルもキーに含める。
    # @ask は同じ質問文なら再実行ごとに外部検索を繰り返さない。
    if name in {"onenote", "命令"} and current_output_file:
        output_part = str(current_output_file.resolve())
    raw = f"{name}\n{body}\n{output_part}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def normalize_task_title(title: str) -> str:
    return re.sub(r"\s+", " ", title.strip()).casefold()


def task_dedupe_key(title: str, tasklist_id: str) -> str:
    raw = f"{tasklist_id}\n{normalize_task_title(title)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def load_known_google_task_keys(output_dir: Path, tasklist_id: str, include_queue: bool = True) -> set[str]:
    keys = set()
    filenames = ["google_todo_synced.jsonl"]
    if include_queue:
        filenames.append("google_todo_queue.jsonl")
    for filename in filenames:
        path = output_dir / filename
        if not path.exists():
            continue
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            if not line.strip():
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            key = record.get("dedupe_key")
            title = str(record.get("title", "")).strip()
            if not key and title:
                key = task_dedupe_key(title, tasklist_id)
            if key:
                keys.add(str(key))
    return keys


def new_command_report() -> dict:
    return {
        "ask_answered": 0,
        "ask_failed": 0,
        "ask_duplicate": 0,
        "google_tasks_new": 0,
        "google_tasks_duplicate": 0,
        "google_tasks_pending": 0,
        "google_tasks_verified": 0,
        "google_tasks_verified_missing": 0,
        "google_tasks_manual": 0,
        "onenote_status": "not_requested",
        "onenote_detail": "",
    }


def merge_command_report(target: dict, source: dict | None) -> dict:
    if not source:
        return target
    for key in (
        "ask_answered",
        "ask_failed",
        "ask_duplicate",
        "google_tasks_new",
        "google_tasks_duplicate",
        "google_tasks_pending",
        "google_tasks_verified",
        "google_tasks_verified_missing",
        "google_tasks_manual",
    ):
        target[key] = int(target.get(key, 0)) + int(source.get(key, 0))
    if source.get("onenote_status") and source.get("onenote_status") != "not_requested":
        target["onenote_status"] = source["onenote_status"]
        target["onenote_detail"] = source.get("onenote_detail", "")
    return target


def print_command_report(report: dict, output_dir: Path) -> None:
    queue_path = output_dir / "google_todo_queue.jsonl"
    pending = report.get("google_tasks_pending", 0)
    if queue_path.exists() and queue_path.stat().st_size == 0:
        pending = 0
    print(
        "Google Tasks: "
        f"新規{report.get('google_tasks_new', 0)}件 / "
        f"重複{report.get('google_tasks_duplicate', 0)}件 / "
        f"未同期{pending}件"
    )
    if report.get("google_tasks_manual"):
        print(f"Google Tasks手動投入待ち: {report.get('google_tasks_manual', 0)}件")
    if report.get("ask_answered") or report.get("ask_failed") or report.get("ask_duplicate"):
        print(
            f"@ask: 回答{report.get('ask_answered', 0)}件 / "
            f"重複{report.get('ask_duplicate', 0)}件 / "
            f"失敗{report.get('ask_failed', 0)}件"
        )
    if report.get("google_tasks_verified") or report.get("google_tasks_verified_missing"):
        print(
            "Google Tasks検証: "
            f"確認済み{report.get('google_tasks_verified', 0)}件 / "
            f"未確認{report.get('google_tasks_verified_missing', 0)}件"
        )
    detail = report.get("onenote_detail", "")
    suffix = f" ({detail})" if detail else ""
    print(f"OneNote: {report.get('onenote_status', 'not_requested')}{suffix}")


def read_manual_todo_copy_items(path: Path | None) -> list[str]:
    if path is None or not path.exists():
        return []
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    match = re.search(r"^## コピー用\s*\n(?P<body>.*?)(?=^## |\Z)", text, flags=re.MULTILINE | re.DOTALL)
    if not match:
        return []
    items: list[str] = []
    for line in match.group("body").splitlines():
        cleaned = line.strip().strip("-*・□☐✅ \t")
        if cleaned:
            items.append(cleaned)
    return items


def audit_todo_registration(text: str, output_dir: Path, manual_todo_path: Path | None = None) -> tuple[int, list[str]]:
    todo_items = []
    for command in parse_at_commands(text):
        if command["name"] == "todo":
            todo_items.extend(split_todo_body(command["body"]))

    manual_items = read_manual_todo_copy_items(manual_todo_path)
    manual_keys = {normalize_task_title(item) for item in manual_items}
    if manual_keys:
        missing_manual = [
            item
            for item in dict.fromkeys(todo_items)
            if normalize_task_title(item) not in manual_keys
        ]
        if not missing_manual:
            return len(todo_items), []

    log_path = output_dir / "command_execution_log.md"
    if not log_path.exists():
        return len(todo_items), list(dict.fromkeys(todo_items))

    log = log_path.read_text(encoding="utf-8-sig")
    seen = set(re.findall(r"## .* @todo Google Tasks登録\n\n- (.*?)\n", log))
    seen.update(re.findall(r"## .* @todo Google Tasks既存タスクのためスキップ\n\n- (.*?)\n", log))
    seen.update(re.findall(r"## .* @todo Google Tasks実体確認OK\n\n- (.*?)\n", log))
    seen_keys = {normalize_task_title(item) for item in seen}
    if manual_keys:
        seen_keys.update(manual_keys)
    missing = [item for item in dict.fromkeys(todo_items) if normalize_task_title(item) not in seen_keys]
    return len(todo_items), missing


def append_markdown_log(output_dir: Path, filename: str, title: str, body: str) -> Path:
    output_dir.mkdir(exist_ok=True)
    path = output_dir / filename
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with path.open("a", encoding="utf-8-sig") as f:
        f.write(f"\n## {timestamp} {title}\n\n")
        f.write(body.rstrip() + "\n")
    return path


def append_jsonl(output_dir: Path, filename: str, record: dict) -> Path:
    output_dir.mkdir(exist_ok=True)
    path = output_dir / filename
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
    return path


def split_todo_body(body: str) -> list[str]:
    todos = []
    for line in body.splitlines():
        if is_command_boundary(line):
            break
        item = line.strip().strip("-*・□☐✅ \t")
        if re.fullmatch(r"[-ーｰ－―\s]+", item):
            break
        if item:
            todos.append(item)
    if not todos and body.strip():
        todos.append(body.strip())
    return todos


def extract_explicit_todo_items(text: str) -> list[str]:
    """Return only tasks written inside @todo/＠todo command blocks."""
    items: list[str] = []
    seen: set[str] = set()
    for command in parse_at_commands(text):
        if command["name"].casefold() != "todo":
            continue
        for item in split_todo_body(command["body"]):
            key = normalize_task_title(item)
            if key and key not in seen:
                seen.add(key)
                items.append(item)
    return items


def build_google_tasks_service(credentials_path: Path, token_path: Path, interactive_auth: bool | None = None):
    """Google Tasks API serviceを作る。依存ライブラリが無い場合は呼び出し側でログ化する。"""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/tasks"]
    creds = None
    if interactive_auth is None:
        interactive_auth = os.getenv("GOOGLE_TASKS_INTERACTIVE_AUTH", "").lower() in {"1", "true", "yes", "on"}

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as exc:
                backup_path = token_path.with_suffix(token_path.suffix + f".revoked_{datetime.now().strftime('%Y%m%d_%H%M%S')}.bak")
                try:
                    token_path.replace(backup_path)
                except OSError:
                    pass
                append_markdown_log(
                    WORKSPACE_DIR / DEFAULT_OUTPUT_DIR,
                    "command_execution_log.md",
                    "@todo Google Tasksトークン再認証へ切替",
                    (
                        f"既存トークンの更新に失敗したため、OAuth再認証に切り替えます。\n\n"
                        f"error: {exc}\n\n"
                        f"backup: `{backup_path}`"
                    ),
                )
                creds = None
        if not creds or not creds.valid:
            if not interactive_auth:
                raise RuntimeError(
                    "Google Tasks OAuth認証が未完了または失効しています。"
                    "自動実行ではブラウザ認証を開きません。"
                    "`powershell -File .\\agent.ps1 google-auth` を先に実行してください。"
                )
            if not credentials_path.exists():
                raise FileNotFoundError(
                    f"Google OAuthクライアントJSONが見つかりません: {credentials_path}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes)
            print("Google Tasks OAuth認証が必要です。ブラウザが開いたら許可してください。", flush=True)
            creds = flow.run_local_server(
                port=0,
                open_browser=True,
                authorization_prompt_message=(
                    "Google Tasks OAuth認証URL:\n{url}\n"
                    "ブラウザが自動で開かない場合は、このURLを開いてください。\n"
                ),
                success_message="Google Tasks OAuth認証が完了しました。このブラウザを閉じてください。",
            )

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("tasks", "v1", credentials=creds)


def google_task_title_exists(service, tasklist_id: str, title: str) -> bool:
    normalized_title = normalize_task_title(title)
    return normalized_title in get_google_task_titles(service, tasklist_id)


def get_google_task_titles(service, tasklist_id: str) -> set[str]:
    titles = set()
    page_token = None
    while True:
        result = (
            service.tasks()
            .list(
                tasklist=tasklist_id,
                maxResults=100,
                pageToken=page_token,
                showCompleted=True,
                showHidden=True,
            )
            .execute()
        )
        for task in result.get("items", []):
            titles.add(normalize_task_title(task.get("title", "")))
        page_token = result.get("nextPageToken")
        if not page_token:
            return titles


def add_google_task(
    title: str,
    output_dir: Path,
    credentials_path: Path,
    token_path: Path,
    tasklist_id: str,
    service=None,
    existing_titles: set[str] | None = None,
) -> tuple[bool, str]:
    try:
        if service is None:
            service = build_google_tasks_service(credentials_path, token_path)
        normalized_title = normalize_task_title(title)
        if existing_titles is None:
            existing_titles = get_google_task_titles(service, tasklist_id)
        if normalized_title in existing_titles:
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@todo Google Tasks既存タスクのためスキップ",
                f"- {title}",
            )
            return True, "duplicate"
        created = (
            service.tasks()
            .insert(tasklist=tasklist_id, body={"title": title})
            .execute()
        )
        existing_titles.add(normalized_title)
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks登録",
            f"- {title}\n\nGoogle task id: `{created.get('id', '')}`",
        )
        return True, "created"
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks登録失敗",
            (
                f"- {title}\n\n"
                f"error: {exc}\n\n"
                "必要な場合は `google-api-python-client`, `google-auth-oauthlib`, "
                "`google-auth-httplib2` をインストールし、"
                "`--google-credentials` にOAuthクライアントJSONを指定してください。"
            ),
        )
        return False, "failed"


def ensure_google_tasks_registered(
    *,
    items: list[str],
    output_dir: Path,
    credentials_path: Path,
    token_path: Path,
    tasklist_id: str,
    service,
) -> tuple[int, list[str]]:
    """Google Tasks本体に@todoが存在することを確認し、不足分は再追加する。"""
    unique_items = list(dict.fromkeys(item.strip() for item in items if item.strip()))
    if not unique_items:
        return 0, []

    try:
        existing_titles = get_google_task_titles(service, tasklist_id)
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks実体確認失敗",
            f"error: {exc}",
        )
        return 0, unique_items

    missing = [
        item for item in unique_items
        if normalize_task_title(item) not in existing_titles
    ]
    if missing:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks実体確認で不足を検出",
            "\n".join(f"- {item}" for item in missing),
        )

    for item in missing:
        synced, _status = add_google_task(
            title=item,
            output_dir=output_dir,
            credentials_path=credentials_path,
            token_path=token_path,
            tasklist_id=tasklist_id,
            service=service,
            existing_titles=existing_titles,
        )
        if not synced:
            append_jsonl(
                output_dir,
                "google_todo_queue.jsonl",
                {
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "title": item,
                    "source": "summarize_note.py @todo final verification",
                    "status": "pending_google_todo_sync_after_verification",
                    "dedupe_key": task_dedupe_key(item, tasklist_id),
                },
            )

    try:
        verified_titles = get_google_task_titles(service, tasklist_id)
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks再確認失敗",
            f"error: {exc}",
        )
        return 0, missing

    still_missing = [
        item for item in unique_items
        if normalize_task_title(item) not in verified_titles
    ]
    verified_count = len(unique_items) - len(still_missing)
    append_markdown_log(
        output_dir,
        "command_execution_log.md",
        "@todo Google Tasks実体確認OK" if not still_missing else "@todo Google Tasks再確認後も未登録",
        "\n".join(f"- {item}" for item in (unique_items if not still_missing else still_missing)),
    )
    return verified_count, still_missing


def find_latest_markdown(output_dir: Path) -> Path | None:
    markdown_files = sorted(
        output_dir.glob("*_日誌.md"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return markdown_files[0] if markdown_files else None


def markdown_to_simple_html(markdown: str, title: str) -> str:
    import html

    body_parts = []
    in_list = False

    def close_list() -> None:
        nonlocal in_list
        if in_list:
            body_parts.append("</ul>")
            in_list = False

    for raw_line in markdown.replace("\r\n", "\n").split("\n"):
        stripped = raw_line.strip()
        if not stripped:
            close_list()
            continue

        heading = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading:
            close_list()
            level = min(len(heading.group(1)), 4)
            body_parts.append(f"<h{level}>{html.escape(heading.group(2))}</h{level}>")
            continue

        bullet = re.match(r"^[-*]\s+(.+)$", stripped)
        if bullet:
            if not in_list:
                body_parts.append("<ul>")
                in_list = True
            body_parts.append(f"<li>{html.escape(bullet.group(1))}</li>")
            continue

        close_list()
        body_parts.append(f"<p>{html.escape(stripped)}</p>")

    close_list()
    return "\n".join(body_parts)


def markdown_to_onenote_page_xml(
    page_title: str,
    markdown: str,
    todo_items: Iterable[str] = (),
) -> str:
    def cdata(text: str) -> str:
        return text.replace("]]>", "]]]]><![CDATA[>")

    oe_parts = []
    for raw_line in markdown.replace("\r\n", "\n").split("\n"):
        stripped = raw_line.strip()
        if not stripped:
            continue
        heading = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading:
            stripped = heading.group(2).strip()
        oe_parts.append(
            "      <one:OE>\n"
            f"        <one:T><![CDATA[{cdata(stripped)}]]></one:T>\n"
            "      </one:OE>"
        )

    if not oe_parts:
        oe_parts.append(
            "      <one:OE>\n"
            "        <one:T><![CDATA[(empty note)]]></one:T>\n"
            "      </one:OE>"
        )

    deduped_todos: list[str] = []
    seen_todos: set[str] = set()
    for item in todo_items:
        cleaned = item.strip().strip("-*・□☐✅ \t")
        key = normalize_task_title(cleaned)
        if cleaned and key and key not in seen_todos:
            seen_todos.add(key)
            deduped_todos.append(cleaned)

    # Ctrl+1 in OneNote applies the built-in To Do tag.  TagDef declares that
    # tag for the page; each final OE refers to it as an unchecked task.
    for item in deduped_todos:
        oe_parts.append(
            "      <one:OE>\n"
            "        <one:Tag index=\"0\" completed=\"false\"/>\n"
            f"        <one:T><![CDATA[{cdata(item)}]]></one:T>\n"
            "      </one:OE>"
        )

    tag_definition = ""
    if deduped_todos:
        tag_definition = (
            "  <one:TagDef index=\"0\" type=\"0\" symbol=\"3\" "
            "fontColor=\"automatic\" highlightColor=\"none\" name=\"To Do\"/>\n"
        )

    return f"""<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="__PAGE_ID__">
{tag_definition}  <one:Title>
    <one:OE>
      <one:T><![CDATA[{cdata(page_title)}]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0"/>
    <one:OEChildren>
{chr(10).join(oe_parts)}
    </one:OEChildren>
  </one:Outline>
</one:Page>
"""


def select_new_onenote_todos(todo_items: Iterable[str], seen_source_keys: Iterable[str]) -> tuple[list[str], list[str]]:
    """Select source Todos not previously sent, independent of later OneNote edits."""
    seen = {key for key in seen_source_keys if key}
    all_keys = list(seen)
    new_items: list[str] = []
    for item in todo_items:
        cleaned = item.strip().strip("-*・□☐✅ \t")
        key = normalize_task_title(cleaned)
        if not cleaned or not key or key in seen:
            continue
        seen.add(key)
        all_keys.append(key)
        new_items.append(cleaned)
    return new_items, all_keys


def load_onenote_todo_source_state(output_dir: Path) -> dict:
    path = output_dir / "onenote_todo_source_state.json"
    if not path.exists():
        return {"pages": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        return data if isinstance(data, dict) else {"pages": {}}
    except (OSError, json.JSONDecodeError):
        return {"pages": {}}


def save_onenote_todo_source_state(output_dir: Path, state: dict) -> None:
    path = output_dir / "onenote_todo_source_state.json"
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_onenote_fallback_html(markdown_file: Path, output_dir: Path) -> Path:
    markdown = markdown_file.read_text(encoding="utf-8-sig")
    title = markdown_file.stem
    html_body = markdown_to_simple_html(markdown, title)
    html_file = output_dir / f"{markdown_file.stem}_onenote.html"
    html_file.write_text(
        "\n".join(
            [
                "<!doctype html>",
                "<html lang=\"ja\">",
                "<head><meta charset=\"utf-8\">",
                f"<title>{title}</title>",
                "<style>body{font-family:Meiryo,'Yu Gothic',sans-serif;line-height:1.7;max-width:920px;margin:32px auto;padding:0 20px;} h1,h2,h3{line-height:1.35;} li{margin:4px 0;}</style>",
                "</head>",
                "<body>",
                f"<h1>{title}</h1>",
                html_body,
                "</body></html>",
            ]
        ),
        encoding="utf-8",
    )
    return html_file


def onenote_page_title(markdown_file: Path, markdown: str) -> str:
    # 同じ日付の日誌はOneNote上の同一タイトルにまとめる。
    # 書き込み時は古いページを削除してから新しいページを作る。
    # unique_output_pathで作られる `_2`, `_3` などの枝番はタイトルから外す。
    return re.sub(r"_\d+$", "", markdown_file.stem)


def add_markdown_to_onenote(
    markdown_file: Path,
    section_name: str,
    output_dir: Path,
    todo_items: Iterable[str] = (),
) -> tuple[bool, str]:
    markdown = markdown_file.read_text(encoding="utf-8-sig")
    title = onenote_page_title(markdown_file, markdown)
    todo_state = load_onenote_todo_source_state(output_dir)
    pages = todo_state.setdefault("pages", {})
    page_state = pages.setdefault(title, {})
    bootstrap_source_state = "seen_source_keys" not in page_state
    seen_source_keys = page_state.get("seen_source_keys", [])
    new_todo_items, updated_source_keys = select_new_onenote_todos(todo_items, seen_source_keys)
    page_xml = markdown_to_onenote_page_xml(title, markdown, todo_items=new_todo_items)
    notebook_name = os.getenv("ONENOTE_NOTEBOOK_NAME", DEFAULT_ONENOTE_NOTEBOOK_NAME)

    ps_script = r"""
param(
  [Parameter(Mandatory=$true)][string]$NotebookName,
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$PageXmlPath,
  [switch]$BootstrapSourceState
)
$ErrorActionPreference = "Stop"
$one = New-Object -ComObject OneNote.Application
$hierarchy = ""
$one.GetHierarchy("", 4, [ref]$hierarchy)
[xml]$doc = $hierarchy
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $doc.DocumentElement.NamespaceURI)
$notebook = $doc.SelectNodes("//one:Notebook", $ns) |
  Where-Object { $_.name -eq $NotebookName } |
  Select-Object -First 1
if ($null -eq $notebook) {
  throw "OneNote notebook not found: $NotebookName"
}
$section = $null
foreach ($candidate in $notebook.SelectNodes(".//one:Section", $ns)) {
  if ($candidate.name -eq $SectionName) {
    $section = $candidate
    break
  }
}
if ($null -eq $section) {
  throw "OneNote section not found: $SectionName"
}

$sectionHierarchy = ""
$one.GetHierarchy($section.ID, 4, [ref]$sectionHierarchy)
[xml]$sectionDoc = $sectionHierarchy
$sectionNs = New-Object System.Xml.XmlNamespaceManager($sectionDoc.NameTable)
$sectionNs.AddNamespace("one", $sectionDoc.DocumentElement.NamespaceURI)
$pageXmlText = [System.IO.File]::ReadAllText($PageXmlPath, [System.Text.Encoding]::UTF8)
[xml]$newPageDoc = $pageXmlText
$newNs = New-Object System.Xml.XmlNamespaceManager($newPageDoc.NameTable)
$newNs.AddNamespace("one", $newPageDoc.DocumentElement.NamespaceURI)

function Get-TaskKey([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
  # Ignore layout-only differences introduced by OCR or manual cleanup.
  return ([regex]::Replace($Text.Trim().ToLowerInvariant(), "[\s,，、。．・]+", ""))
}

# Pull the newly generated task nodes out temporarily. They are appended again
# only after every existing OneNote task, so an update never rewrites progress.
$incomingTasks = @()
$incomingNodes = @($newPageDoc.SelectNodes("//one:Outline//one:OE[one:Tag]", $newNs))
foreach ($node in $incomingNodes) {
  $textNode = $node.SelectSingleNode("./one:T", $newNs)
  if ($null -ne $textNode -and -not [string]::IsNullOrWhiteSpace($textNode.InnerText)) {
    $incomingTasks += [pscustomobject]@{
      Text = $textNode.InnerText.Trim()
      Completed = $false
    }
  }
  [void]$node.ParentNode.RemoveChild($node)
}

# The existing OneNote list is the source of truth. Preserve its text, order,
# and checked state exactly; only unseen incoming tasks are appended later.
$existingTasks = @()
$pageId = $null
$deletedCount = 0
foreach ($page in $sectionDoc.SelectNodes("//one:Page", $sectionNs)) {
  if ($page.name -eq $PageTitle) {
    $existingPageText = ""
    $one.GetPageContent($page.ID, [ref]$existingPageText, 2)
    [xml]$existingPageDoc = $existingPageText
    $existingNs = New-Object System.Xml.XmlNamespaceManager($existingPageDoc.NameTable)
    $existingNs.AddNamespace("one", $existingPageDoc.DocumentElement.NamespaceURI)
    $todoIndexes = @{}
    foreach ($definition in $existingPageDoc.SelectNodes("//one:TagDef", $existingNs)) {
      if ([string]$definition.name -match "^(To Do|Todo|タスク)") {
        $todoIndexes[[string]$definition.index] = $true
      }
    }
    foreach ($oe in $existingPageDoc.SelectNodes("//one:Outline//one:OE[one:Tag]", $existingNs)) {
      $tag = $oe.SelectSingleNode("./one:Tag", $existingNs)
      if (-not $todoIndexes.ContainsKey([string]$tag.index)) { continue }
      $textNode = $oe.SelectSingleNode("./one:T", $existingNs)
      if ($null -eq $textNode -or [string]::IsNullOrWhiteSpace($textNode.InnerText)) { continue }
      $existingTasks += [pscustomobject]@{
        Text = $textNode.InnerText.Trim()
        Completed = (([string]$tag.completed).ToLowerInvariant() -eq "true")
      }
    }
    $one.DeleteHierarchy($page.ID)
    $deletedCount += 1
  }
}
if ($deletedCount -gt 0) {
  Start-Sleep -Milliseconds 800
}

$pageId = ""
$one.CreateNewPage($section.ID, [ref]$pageId, 0)
if ($deletedCount -gt 0) {
  $action = "recreated_deleted_$deletedCount"
} else {
  $action = "created"
}

$effectiveIncomingTasks = $incomingTasks
if ($BootstrapSourceState -and $existingTasks.Count -gt 0) {
  # Existing OneNote text may contain user corrections that are more reliable
  # than OCR/rawtext. On first state migration, keep OneNote untouched and only
  # let Python record the already-seen source keys after this succeeds.
  $effectiveIncomingTasks = @()
}

$combinedTasks = @()
$seenTaskKeys = @{}
foreach ($task in $existingTasks) {
  $key = Get-TaskKey $task.Text
  if ($key -and -not $seenTaskKeys.ContainsKey($key)) {
    $seenTaskKeys[$key] = $true
    $combinedTasks += $task
  }
}
$appendedCount = 0
foreach ($task in $effectiveIncomingTasks) {
  $key = Get-TaskKey $task.Text
  if ($key -and -not $seenTaskKeys.ContainsKey($key)) {
    $seenTaskKeys[$key] = $true
    $combinedTasks += $task
    $appendedCount += 1
  }
}

if ($combinedTasks.Count -gt 0) {
  $tagDefinition = $newPageDoc.SelectSingleNode("//one:TagDef[@name='To Do']", $newNs)
  if ($null -eq $tagDefinition) {
    $tagDefinition = $newPageDoc.CreateElement("one", "TagDef", $newPageDoc.DocumentElement.NamespaceURI)
    $tagDefinition.SetAttribute("index", "0")
    $tagDefinition.SetAttribute("type", "0")
    $tagDefinition.SetAttribute("symbol", "3")
    $tagDefinition.SetAttribute("fontColor", "automatic")
    $tagDefinition.SetAttribute("highlightColor", "none")
    $tagDefinition.SetAttribute("name", "To Do")
    $titleNode = $newPageDoc.SelectSingleNode("//one:Title", $newNs)
    [void]$newPageDoc.DocumentElement.InsertBefore($tagDefinition, $titleNode)
  }
  $taskContainer = $newPageDoc.SelectSingleNode("//one:Outline/one:OEChildren", $newNs)
  foreach ($task in $combinedTasks) {
    $oe = $newPageDoc.CreateElement("one", "OE", $newPageDoc.DocumentElement.NamespaceURI)
    $tag = $newPageDoc.CreateElement("one", "Tag", $newPageDoc.DocumentElement.NamespaceURI)
    $tag.SetAttribute("index", "0")
    $tag.SetAttribute("completed", $(if ($task.Completed) { "true" } else { "false" }))
    [void]$oe.AppendChild($tag)
    $text = $newPageDoc.CreateElement("one", "T", $newPageDoc.DocumentElement.NamespaceURI)
    [void]$text.AppendChild($newPageDoc.CreateCDataSection([string]$task.Text))
    [void]$oe.AppendChild($text)
    [void]$taskContainer.AppendChild($oe)
  }
}

$pageXml = $newPageDoc.OuterXml
$pageXml = $pageXml.Replace("__PAGE_ID__", $pageId)
$one.UpdatePageContent($pageXml)
Write-Output "$action $pageId existing_tasks=$($existingTasks.Count) incoming_tasks=$($incomingTasks.Count) appended_tasks=$appendedCount total_tasks=$($combinedTasks.Count) bootstrap=$BootstrapSourceState"
"""

    with tempfile.TemporaryDirectory(prefix="summarize_note5_onenote_") as tmp:
        tmp_dir = Path(tmp)
        ps_path = tmp_dir / "add_to_onenote.ps1"
        xml_path = tmp_dir / "page.xml"
        ps_path.write_text(ps_script, encoding="utf-8")
        xml_path.write_text(page_xml, encoding="utf-8")
        command = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ps_path),
                "-NotebookName",
                notebook_name,
                "-SectionName",
                section_name,
                "-PageTitle",
                title,
                "-PageXmlPath",
                str(xml_path),
            ]
        if bootstrap_source_state:
            command.append("-BootstrapSourceState")
        result = subprocess.run(
            command,
            cwd=str(WORKSPACE_DIR),
            text=True,
            capture_output=True,
            timeout=120,
            check=False,
        )

    if result.returncode == 0:
        page_state["seen_source_keys"] = updated_source_keys
        page_state["initialized_from_onenote"] = bootstrap_source_state
        page_state["updated_at"] = datetime.now().isoformat(timespec="seconds")
        save_onenote_todo_source_state(output_dir, todo_state)
        return True, result.stdout.strip()

    html_file = write_onenote_fallback_html(markdown_file, output_dir)
    return False, (
        result.stderr.strip()
        or result.stdout.strip()
        or f"OneNote COM failed. Fallback HTML created: {html_file}"
    )


def run_onenote_command(
    body: str,
    output_dir: Path,
    section_name: str,
    current_output_file: Path | None,
    todo_items: Iterable[str] = (),
) -> tuple[str, str]:
    markdown_file = current_output_file if current_output_file else find_latest_markdown(output_dir)
    if not markdown_file or not markdown_file.exists():
        queue_deferred_command("onenote", body, output_dir)
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@onenote 実行不可",
            "追加対象のMarkdown日誌が見つかりませんでした。",
        )
        return "failed", "markdown not found"

    try:
        ok, detail = add_markdown_to_onenote(
            markdown_file,
            section_name,
            output_dir,
            todo_items=todo_items,
        )
        if ok:
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@onenote OneNote作成/更新",
                f"markdown: `{markdown_file}`\n\nsection: `{section_name}`\n\npage id: `{detail}`",
            )
            action = detail.split(maxsplit=1)[0] if detail else "ok"
            return action, detail
        else:
            queue_deferred_command("onenote", body, output_dir)
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@onenote フォールバック",
                f"markdown: `{markdown_file}`\n\nsection: `{section_name}`\n\n{detail}",
            )
            return "fallback", detail
    except Exception as exc:
        queue_deferred_command("onenote", body, output_dir)
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@onenote 実行エラー",
            f"markdown: `{markdown_file}`\n\nsection: `{section_name}`\n\nerror: {exc}",
        )
        return "failed", str(exc)


def resolve_safe_python_command(command_text: str) -> list[str]:
    parts = shlex.split(command_text, posix=False)
    if not parts:
        raise ValueError("@python の実行内容が空です")

    script = Path(parts[0].strip('"'))
    if not script.is_absolute():
        script = WORKSPACE_DIR / script
    script = script.resolve()

    workspace = WORKSPACE_DIR.resolve()
    try:
        script.relative_to(workspace)
    except ValueError as exc:
        raise ValueError(f"作業フォルダ外の.pyは実行しません: {script}") from exc

    if script.suffix.lower() != ".py":
        raise ValueError(f".pyファイルだけ実行できます: {script.name}")
    if not script.exists():
        raise FileNotFoundError(f"Pythonファイルが見つかりません: {script}")

    return [sys.executable, str(script), *parts[1:]]


def run_python_command(command_text: str, output_dir: Path, timeout: int) -> None:
    try:
        cmd = resolve_safe_python_command(command_text)
        result = subprocess.run(
            cmd,
            cwd=str(WORKSPACE_DIR),
            text=True,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        body = [
            f"**command:** `{command_text}`",
            f"**returncode:** `{result.returncode}`",
            "",
            "### stdout",
            "```text",
            result.stdout.strip(),
            "```",
            "",
            "### stderr",
            "```text",
            result.stderr.strip(),
            "```",
        ]
        append_markdown_log(output_dir, "command_execution_log.md", "@python 実行結果", "\n".join(body))
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@python 実行エラー",
            f"**command:** `{command_text}`\n\n**error:** {exc}",
        )


def append_ask_answer_to_journal(markdown_file: Path, question: str, answer: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with markdown_file.open("a", encoding="utf-8-sig") as f:
        f.write("\n\n---\n\n")
        f.write(f"## @ask 回答 ({timestamp})\n\n")
        f.write("### 質問\n\n")
        f.write(question.strip() + "\n\n")
        f.write("### 回答\n\n")
        f.write(answer.strip() + "\n")


def ask_answer_already_logged(output_dir: Path, question: str) -> bool:
    analysis_path = output_dir / "analysis.md"
    if not analysis_path.exists():
        return False
    try:
        text = analysis_path.read_text(encoding="utf-8-sig", errors="replace")
    except Exception:
        return False
    return question.strip() in text and "@ask 解析結果" in text


def create_ask_response(client: Anthropic, model: str, prompt: str) -> str:
    timeout_seconds = float(os.getenv("ASK_CLAUDE_TIMEOUT_SECONDS", "90"))
    return llm_client.create_response(client, model, prompt, timeout=timeout_seconds, web_search=True)


def run_ask_command(
    question: str,
    output_dir: Path,
    client: Anthropic | None,
    model: str,
    current_output_file: Path | None = None,
) -> bool:
    # --commands-only の場合でも @ask が使えるように、必要になった時点で client を作る。
    if client is None:
        try:
            client = llm_client.get_client()
        except Exception as exc:
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@ask 実行エラー",
                f"Anthropic client を作成できませんでした。ANTHROPIC_API_KEY を確認してください。\n\nerror: {exc}",
            )
            return False

    prompt = f"""
あなたは研究者の日誌メモに含まれる質問へ答えるAI秘書です。
今日の日付は {date.today().isoformat()} です。

以下の質問・メモに対して、日本語で回答してください。
時事性がある質問、製品発売時期、価格、予定、ニュース、法律・制度、論文情報など、
現在性が重要な内容は、利用可能ならWeb検索を使って最新情報を確認してください。
確認できた情報、未確定の情報、次に見るべき公式情報を分けてください。
根拠として使った公式ページ・信頼できる情報源があれば、URLまたは出典名も書いてください。
調べきれない場合は、断定せず「未確認」と明記してください。

質問・メモ:
{question}
"""
    try:
        answer = create_ask_response(client, model, prompt)
        append_markdown_log(
            output_dir,
            "analysis.md",
            "@ask 解析結果",
            f"### 質問\n\n{question}\n\n### 回答\n\n{answer}",
        )
        if current_output_file:
            append_ask_answer_to_journal(current_output_file, question, answer)
        return True
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@ask 実行エラー",
            f"### 質問\n\n{question}\n\n### エラー\n\n{exc}",
        )
        return False




def queue_deferred_command(name: str, body: str, output_dir: Path) -> None:
    """未実装または外部連携が必要な @ コマンドを安全にキューへ保存する。"""
    timestamp = datetime.now().isoformat(timespec="seconds")
    description = SUPPORTED_AT_COMMANDS.get(name, "未定義コマンド")
    record = {
        "created_at": timestamp,
        "command": name,
        "description": description,
        "body": body,
        "status": "queued",
        "source": "summarize_note5.py",
    }
    queue_file = QUEUE_FILE_BY_COMMAND.get(name, "unknown_command_queue.jsonl")
    markdown_file = MARKDOWN_LOG_BY_COMMAND.get(name, "unknown_commands.md")
    append_jsonl(output_dir, queue_file, record)
    append_markdown_log(
        output_dir,
        markdown_file,
        f"@{name} {description}",
        body,
    )

def execute_at_commands(
    text: str,
    output_dir: Path,
    state: dict,
    client: Anthropic,
    model: str,
    python_timeout: int,
    sync_google_todos: bool,
    google_credentials: Path,
    google_token: Path,
    google_tasklist: str,
    onenote_section: str,
    current_output_file: Path | None = None,
    skip_ask: bool = False,
    manual_google_tasks: bool = False,
    onenote_todo_items: Iterable[str] = (),
    depth: int = 0,
) -> tuple[dict, dict]:
    report = new_command_report()
    if depth > 5:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@命令 実行停止",
            "命令の入れ子が深すぎるため停止しました。",
        )
        return state, report

    commands = parse_at_commands(text)
    if not commands:
        return state, report

    executed_hashes = set(state.get("executed_command_hashes", []))
    known_task_keys = set(state.get("known_google_task_keys", []))
    known_task_keys.update(state.get("synced_google_task_keys", []))
    known_task_keys.update(load_known_google_task_keys(output_dir, google_tasklist))
    new_hashes = []
    new_task_keys = []
    all_todo_items: list[str] = []
    google_service = None
    existing_google_titles = None
    if sync_google_todos and any(command["name"] == "todo" for command in commands):
        try:
            google_service = build_google_tasks_service(google_credentials, google_token)
            existing_google_titles = get_google_task_titles(google_service, google_tasklist)
        except Exception as exc:
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@todo Google Tasks初期化失敗",
                f"error: {exc}",
            )
            sync_google_todos = False

    for command in commands:
        name = command["name"]
        body = command["body"].strip()
        h = command_hash(command, current_output_file=current_output_file)
        if h in executed_hashes and name != "todo":
            continue

        timestamp = datetime.now().isoformat(timespec="seconds")

        if name == "todo":
            processed_items = []
            skipped_items = []
            todo_items = split_todo_body(body)
            all_todo_items.extend(todo_items)
            for item_index, item in enumerate(todo_items, start=1):
                item_key = task_dedupe_key(item, google_tasklist)
                if manual_google_tasks:
                    report["google_tasks_manual"] += 1
                    append_candidates(output_dir, "todo_candidates.md", "@todo 手動投入候補", [item])
                    processed_items.append(item)
                    continue
                if item_key in known_task_keys:
                    if (
                        sync_google_todos
                        and existing_google_titles is not None
                        and normalize_task_title(item) not in existing_google_titles
                    ):
                        append_markdown_log(
                            output_dir,
                            "command_execution_log.md",
                            "@todo ローカル重複ログはあるがGoogle Tasks本体にないため再登録",
                            f"- {item}",
                        )
                    elif (
                        sync_google_todos
                        and existing_google_titles is not None
                        and normalize_task_title(item) in existing_google_titles
                    ):
                        report["google_tasks_duplicate"] += 1
                        skipped_items.append(item)
                        append_markdown_log(
                            output_dir,
                            "command_execution_log.md",
                            "@todo Google Tasks既存タスクのためスキップ",
                            f"- {item}",
                        )
                        continue
                    else:
                        report["google_tasks_pending"] += 1
                        synced = False
                        record = {
                            "created_at": timestamp,
                            "title": item,
                            "source": "summarize_note.py @todo",
                            "status": "pending_google_todo_sync_unverified_local_key",
                            "dedupe_key": item_key,
                        }
                        append_jsonl(output_dir, "google_todo_queue.jsonl", record)
                        append_candidates(output_dir, "todo_candidates.md", "@todo 追加候補", [item])
                        processed_items.append(item)
                        append_markdown_log(
                            output_dir,
                            "command_execution_log.md",
                            "@todo Google Tasks未確認のため未同期キューへ戻す",
                            (
                                f"- {item}\n\n"
                                "Google Tasks実体確認ができない状態では、"
                                "ローカル重複ログだけを根拠に同期済み扱いしません。"
                            ),
                        )
                        continue

                synced = False
                if sync_google_todos:
                    synced, sync_status = add_google_task(
                        title=item,
                        output_dir=output_dir,
                        credentials_path=google_credentials,
                        token_path=google_token,
                        tasklist_id=google_tasklist,
                        service=google_service,
                        existing_titles=existing_google_titles,
                    )
                    if synced:
                        new_task_keys.append(item_key)
                        known_task_keys.add(item_key)
                        if sync_status == "duplicate":
                            report["google_tasks_duplicate"] += 1
                        else:
                            report["google_tasks_new"] += 1
                    else:
                        report["google_tasks_pending"] += 1
                else:
                    report["google_tasks_pending"] += 1
                record = {
                    "created_at": timestamp,
                    "title": item,
                    "source": "summarize_note.py @todo",
                    "status": "synced_google_tasks" if synced else "pending_google_todo_sync",
                    "dedupe_key": item_key,
                }
                if not synced:
                    append_jsonl(output_dir, "google_todo_queue.jsonl", record)
                append_candidates(output_dir, "todo_candidates.md", "@todo 追加候補", [item])
                processed_items.append(item)
                if (
                    sync_google_todos
                    and item_index % GOOGLE_TASK_BATCH_SIZE == 0
                    and item_index < len(todo_items)
                ):
                    append_markdown_log(
                        output_dir,
                        "command_execution_log.md",
                        "@todo Google Tasksバッチ区切り",
                        (
                            f"{item_index}件まで処理しました。"
                            f"{GOOGLE_TASK_BATCH_PAUSE_SECONDS:g}秒待ってから続行します。"
                        ),
                    )
                    time.sleep(GOOGLE_TASK_BATCH_PAUSE_SECONDS)
            if processed_items:
                append_markdown_log(output_dir, "command_execution_log.md", "@todo 処理", "\n".join(processed_items))
            if skipped_items and not processed_items:
                append_markdown_log(
                    output_dir,
                    "command_execution_log.md",
                    "@todo 処理",
                    "すべて登録済みだったため、Google Tasksへの新規登録はありませんでした。",
                )

        elif name == "python":
            run_python_command(body, output_dir, timeout=python_timeout)

        elif name == "ask":
            if skip_ask:
                queue_deferred_command("ask", body, output_dir)
                append_markdown_log(
                    output_dir,
                    "command_execution_log.md",
                    "@ask ローカル日誌モードのためスキップ",
                    (
                        "ローカル日誌モードでは外部APIへ質問を送らないため、"
                        "以下の @ask はキュー/ログに保存しました。\n\n"
                        f"### 質問\n\n{body}"
                    ),
                )
                report["ask_skipped"] = int(report.get("ask_skipped", 0)) + 1
            elif ask_answer_already_logged(output_dir, body):
                append_markdown_log(
                    output_dir,
                    "command_execution_log.md",
                    "@ask 重複スキップ",
                    f"### 質問\n\n{body}\n\n既に analysis.md に回答があるため、外部検索を再実行しませんでした。",
                )
                report["ask_duplicate"] = int(report.get("ask_duplicate", 0)) + 1
            else:
                ok = run_ask_command(
                    body,
                    output_dir,
                    client=client,
                    model=model,
                    current_output_file=current_output_file,
                )
                if ok:
                    report["ask_answered"] += 1
                else:
                    report["ask_failed"] += 1

        elif name == "onenote":
            onenote_status, onenote_detail = run_onenote_command(
                body=body,
                output_dir=output_dir,
                section_name=onenote_section,
                current_output_file=current_output_file,
                todo_items=onenote_todo_items,
            )
            report["onenote_status"] = onenote_status
            report["onenote_detail"] = onenote_detail

        elif name == "命令":
            normalized_body = normalize_instruction_body(body)
            nested_commands = parse_at_commands(normalized_body)
            if nested_commands:
                append_markdown_log(
                    output_dir,
                    "command_execution_log.md",
                    "@命令 展開",
                    normalized_body,
                )
                state, nested_report = execute_at_commands(
                    text=normalized_body,
                    output_dir=output_dir,
                    state=state,
                    client=client,
                    model=model,
                    python_timeout=python_timeout,
                    sync_google_todos=sync_google_todos,
                    google_credentials=google_credentials,
                    google_token=google_token,
                    google_tasklist=google_tasklist,
                    onenote_section=onenote_section,
                    current_output_file=current_output_file,
                    skip_ask=skip_ask,
                    manual_google_tasks=manual_google_tasks,
                    onenote_todo_items=onenote_todo_items,
                    depth=depth + 1,
                )
                merge_command_report(report, nested_report)
            else:
                append_markdown_log(
                    output_dir,
                    "command_execution_log.md",
                    "@命令 実行不可",
                    (
                        "本文から実行可能な命令を判定できませんでした。\n\n"
                        "使える形式の例:\n\n"
                        "```text\n"
                        "＠命令\n"
                        "todo STM測定継続\n"
                        "ask 360Hzノイズについて考察して\n"
                        "python daily_paper_search.py\n"
                        "```"
                    ),
                )

        elif name in QUEUE_FILE_BY_COMMAND:
            # @paper, @news, @mail, @onenote, @experiment, @idea,
            # @memo, @reference, @schedule はまずキューに保存する。
            # 実際の外部連携は専用スクリプト側で行う。
            queue_deferred_command(name, body, output_dir)
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                f"@{name} キュー保存",
                body,
            )

        else:
            append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "未定義コマンド",
                f"**command:** `@{name}`\n\n{body}",
            )

        new_hashes.append(h)

    if sync_google_todos and google_service is not None and all_todo_items:
        verified_count, still_missing = ensure_google_tasks_registered(
            items=all_todo_items,
            output_dir=output_dir,
            credentials_path=google_credentials,
            token_path=google_token,
            tasklist_id=google_tasklist,
            service=google_service,
        )
        report["google_tasks_verified"] += verified_count
        report["google_tasks_verified_missing"] += len(still_missing)
        report["google_tasks_pending"] += len(still_missing)

    if new_hashes or new_task_keys:
        state = {
            **state,
            "executed_command_hashes": sorted(executed_hashes.union(new_hashes))[-1000:],
            "known_google_task_keys": sorted(known_task_keys)[-5000:],
            "last_command_processed_at": datetime.now().isoformat(timespec="seconds"),
        }
    return state, report


def append_candidates(output_dir: Path, filename: str, title: str, items: Iterable[str]) -> None:
    items = [item.strip() for item in items if item.strip()]
    if not items:
        return
    output_dir.mkdir(exist_ok=True)
    path = output_dir / filename
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    body = [f"\n## {timestamp} {title}"]
    body.extend(f"- {item}" for item in items)
    path.open("a", encoding="utf-8-sig").write("\n".join(body) + "\n")


TODO_CATEGORIES = (
    ("時刻あり", ("時", ":", "：")),
    ("実験", ("反応", "合成", "NMR", "nmr", "Cage", "水酸化ナトリウム", "処理")),
    ("測定・解析", ("STM", "Raman", "測定", "SEM", "TEM", "解析")),
    ("買う・受け取る", ("受取", "購入", "発注", "買", "在庫")),
    ("事務", ("会議", "会計", "出納", "WEB", "登録", "調査", "記入")),
    ("記録・整理", ("Note", "ノート", "日誌", "メモ", "PPT", "記録", "掃除", "整理")),
)


def todo_has_time(item: str) -> bool:
    return re.search(r"\b\d{1,2}\s*[:：]\s*\d{2}\b", item) is not None or re.search(r"\d{1,2}\s*時", item) is not None


def classify_todo(item: str) -> str:
    if todo_has_time(item):
        return "時刻あり"
    for category, keywords in TODO_CATEGORIES[1:]:
        if any(keyword.casefold() in item.casefold() for keyword in keywords):
            return category
    return "その他"


def grouped_todos(items: Iterable[str]) -> dict[str, list[str]]:
    groups = {category: [] for category, _keywords in TODO_CATEGORIES}
    groups["その他"] = []
    for item in items:
        groups[classify_todo(item)].append(item)
    return {category: values for category, values in groups.items() if values}


def find_todo_variation_notes(items: list[str]) -> list[str]:
    notes: list[str] = []
    joined = "\n".join(items)
    variation_pairs = [
        ("液体チッツ", "液体チッソ"),
        ("アセニト", "アセトニトリル"),
        ("Onenote", "OneNote"),
        ("OutPut", "Output"),
    ]
    for left, right in variation_pairs:
        if left in joined and right in joined:
            notes.append(f"`{left}` と `{right}` が混在しています。表記を確認してください。")
        elif left in joined:
            notes.append(f"`{left}` は表記ゆれ/誤変換の可能性があります。必要なら `{right}` に直してください。")
    normalized_map: dict[str, list[str]] = {}
    for item in items:
        key = re.sub(r"[、。，．\s]+", "", item).casefold()
        key = key.replace("チッツ", "チッソ").replace("onenote", "oneNote".casefold())
        normalized_map.setdefault(key, []).append(item)
    for variants in normalized_map.values():
        unique_variants = list(dict.fromkeys(variants))
        if len(unique_variants) > 1:
            notes.append("似たTODOがあります: " + " / ".join(unique_variants[:4]))
    return list(dict.fromkeys(notes))


def write_manual_todo_list(output_dir: Path, journal_date: str, todo_items: Iterable[str]) -> Path | None:
    items = []
    seen = set()
    for item in todo_items:
        cleaned = item.strip().strip("-*・□☐✅ \t")
        if not cleaned or re.fullmatch(r"[-ーｰ－―\s]+", cleaned):
            continue
        key = normalize_task_title(cleaned)
        if key in seen:
            continue
        seen.add(key)
        items.append(cleaned)
    if not items:
        return None

    groups = grouped_todos(items)
    time_items = [item for item in items if todo_has_time(item)]
    variation_notes = find_todo_variation_notes(items)
    classified_lines: list[str] = []
    for category, grouped_items in groups.items():
        classified_lines.extend([f"### {category}", ""])
        classified_lines.extend(grouped_items)
        classified_lines.append("")
    time_lines = time_items if time_items else ["時刻を含むTODOは検出されませんでした。"]
    variation_lines = [f"- {note}" for note in variation_notes] if variation_notes else ["- 表記ゆれ候補は検出されませんでした。"]

    output_dir.mkdir(exist_ok=True)
    path = output_dir / f"{journal_date}_GoogleTasks手動投入用.md"
    lines = [
        f"# Google Tasks 手動投入用 {journal_date}",
        "",
        f"- Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "- Status: 手動投入待ち",
        "",
        "## コピー用",
        "",
        *items,
        "",
        "## 分類別コピー用",
        "",
        *classified_lines,
        "## 時刻あり",
        "",
        *time_lines,
        "",
        "## 表記ゆれ確認",
        "",
        *variation_lines,
        "",
        "## チェック用",
        "",
        *[f"- [ ] {item}" for item in items],
        "",
        "## 運用メモ",
        "",
        "- Codexはこのファイルを作るだけで、Google Tasksへ自動送信しません。",
        "- Google Tasksへ入れたら、チェック用の `[ ]` を `[x]` に変えるか、日誌に `@taskdone` として記録してください。",
    ]
    content = "\n".join(lines) + "\n"
    path.write_text(content, encoding="utf-8")
    latest_path = output_dir / "GoogleTasks手動投入用.md"
    latest_path.write_text(content, encoding="utf-8")
    return path


def unique_output_path(output_dir: Path, journal_date: str, overwrite: bool) -> Path:
    output_file = output_dir / f"{journal_date}_日誌.md"
    if overwrite or not output_file.exists():
        return output_file

    index = 2
    while True:
        candidate = output_dir / f"{journal_date}_日誌_{index}.md"
        if not candidate.exists():
            return candidate
        index += 1


def write_summary(
    output_dir: Path,
    journal_date: str,
    summary: str,
    overwrite: bool,
    append_output: bool,
) -> Path:
    output_dir.mkdir(exist_ok=True)

    if append_output:
        output_file = output_dir / f"{journal_date}_日誌.md"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with output_file.open("a", encoding="utf-8-sig") as f:
            f.write(f"\n\n---\n\n<!-- auto update: {timestamp} -->\n\n")
            f.write(summary.strip() + "\n")
        return output_file

    output_file = unique_output_path(output_dir, journal_date, overwrite)
    output_file.write_text(summary.strip() + "\n", encoding="utf-8-sig")
    return output_file


def build_command_report_text(
    *,
    output_file: Path | None,
    command_report: dict,
    total_todos: int | None = None,
    missing_todos: list[str] | None = None,
    manual_todo_path: Path | None = None,
) -> str:
    lines: list[str] = []
    if output_file is not None:
        lines.append(f"保存しました: {output_file}")
    if manual_todo_path is not None:
        lines.append(f"Google Tasks手動投入用: {manual_todo_path}")
    pending = int(command_report.get("google_tasks_pending", 0))
    lines.append(
        "Google Tasks: "
        f"新規{command_report.get('google_tasks_new', 0)}件 / "
        f"重複{command_report.get('google_tasks_duplicate', 0)}件 / "
        f"未同期{pending}件"
    )
    if command_report.get("ask_answered") or command_report.get("ask_failed") or command_report.get("ask_duplicate"):
        lines.append(
            f"@ask: 回答{command_report.get('ask_answered', 0)}件 / "
            f"重複{command_report.get('ask_duplicate', 0)}件 / "
            f"失敗{command_report.get('ask_failed', 0)}件"
        )
    if command_report.get("google_tasks_verified") or command_report.get("google_tasks_verified_missing"):
        lines.append(
            "Google Tasks検証: "
            f"確認済み{command_report.get('google_tasks_verified', 0)}件 / "
            f"未確認{command_report.get('google_tasks_verified_missing', 0)}件"
        )
    detail = command_report.get("onenote_detail", "")
    suffix = f" ({detail})" if detail else ""
    lines.append(f"OneNote: {command_report.get('onenote_status', 'not_requested')}{suffix}")
    if total_todos is not None:
        missing_count = len(missing_todos or [])
        lines.append(f"@todo監査: 対象{total_todos}件 / 未確認{missing_count}件")
        lines.extend(f"  - {item}" for item in (missing_todos or []))
    return "\n".join(lines)


def run_verification_agent(
    *,
    request: str,
    command: list[str],
    output: str,
    exit_code: int = 0,
) -> bool:
    try:
        from verification_agent import verify_and_log

        result = verify_and_log(
            agent="日誌Agent",
            request=request,
            command=command,
            output=output,
            exit_code=exit_code,
        )
        print("\n--- 検証Agent ---")
        print(result.to_text())
        return result.status != "failed"
    except Exception as exc:
        print("\n--- 検証Agent ---")
        print(f"検証Agentの実行に失敗しました: {exc}")
        return False


def summarize_once(args: argparse.Namespace, client: Anthropic, reason: str = "manual run") -> bool:
    input_path = resolve_path(args.input)
    output_dir = resolve_path(args.output_dir)
    state_dir = resolve_path(args.state_dir)

    try:
        note = read_note_text(input_path)
    except Exception as exc:
        print(f"[skip] メモを読めませんでした: {exc}")
        return False

    if not note.strip():
        print(f"[skip] 入力メモが空です: {input_path}")
        return False

    state = load_state(state_dir, input_path)
    current_hash = hashlib.sha256(note.encode("utf-8")).hexdigest()
    if state.get("full_hash") == current_hash and args.delta_only and not args.force:
        print(f"[skip] 更新なし: {datetime.now().strftime('%H:%M:%S')}")
        return False

    if args.force:
        target_note = note.strip()
        new_state = {
            **state,
            "processed_length": len(note),
            "full_hash": current_hash,
            "processed_prefix_hash": current_hash,
            "previous_hash": state.get("full_hash"),
            "last_processed_at": datetime.now().isoformat(timespec="seconds"),
        }
    elif args.delta_only:
        target_note, new_state = get_delta_text(note, state)
    else:
        target_note = note.strip()
        new_state = {
            **state,
            "processed_length": len(note),
            "full_hash": current_hash,
            "last_processed_at": datetime.now().isoformat(timespec="seconds"),
        }

    if not target_note.strip():
        save_state(state_dir, input_path, new_state)
        print(f"[skip] 差分なし: {datetime.now().strftime('%H:%M:%S')}")
        return False

    command_target_note = get_delta_command_text(note, state, target_note) if args.delta_only else target_note

    print(f"[run] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} / reason={reason}")
    command_report = new_command_report()

    if args.commands_only:
        new_state, command_report = execute_at_commands(
            text=command_target_note,
            output_dir=output_dir,
            state=new_state,
            client=client,
            model=args.model,
            python_timeout=args.python_timeout,
            sync_google_todos=args.sync_google_todos,
            google_credentials=resolve_path(args.google_credentials),
            google_token=resolve_path(args.google_token),
            google_tasklist=args.google_tasklist,
            onenote_section=args.onenote_section,
            current_output_file=None,
            skip_ask=args.skip_ask,
            manual_google_tasks=args.manual_google_tasks,
        )
        save_state(state_dir, input_path, new_state)
        print("[ok] コマンドだけ実行しました。")
        print_command_report(command_report, output_dir)
        total_todos, missing_todos = audit_todo_registration(command_target_note, output_dir)
        print(f"@todo監査: 対象{total_todos}件 / 未確認{len(missing_todos)}件")
        for item in missing_todos:
            print(f"  - {item}")
        verification_output = build_command_report_text(
            output_file=None,
            command_report=command_report,
            total_todos=total_todos,
            missing_todos=missing_todos,
            manual_todo_path=None,
        )
        run_verification_agent(
            request=f"{args.date} @コマンド実行のみ",
            command=sys.argv,
            output="[ok] コマンドだけ実行しました。\n" + verification_output,
            exit_code=0,
        )
        return True

    if args.local_summary:
        summary_body = local_extractive_summary(target_note, args.date)
    else:
        summary_body = llm_client.create_response(
            client,
            args.model,
            build_prompt(target_note, args.date, delta_only=args.delta_only),
            timeout=float(os.getenv("SUMMARY_CLAUDE_TIMEOUT_SECONDS", "180")),
        )
    summary = summary_body + build_raw_text_appendix(input_path, target_note)

    output_file = write_summary(
        output_dir=output_dir,
        journal_date=args.date,
        summary=summary,
        overwrite=args.overwrite,
        append_output=args.append_output,
    )

    at_commands = parse_at_commands(command_target_note)
    explicit_todo_items = extract_explicit_todo_items(command_target_note)
    manual_todo_items = extract_todo_candidates(command_target_note)
    for command in at_commands:
        if command["name"] == "todo":
            manual_todo_items.extend(split_todo_body(command["body"]))
    manual_todo_path = write_manual_todo_list(output_dir, args.date, manual_todo_items)

    if args.extract_commands:
        block_commands = extract_command_blocks(command_target_note)
        at_command_lines = [f"@{c['name']}\n{c['body']}" for c in at_commands]
        append_candidates(
            output_dir,
            "command_candidates.md",
            "命令候補",
            [*block_commands, *at_command_lines],
        )

    if args.extract_todos:
        append_candidates(
            output_dir,
            "todo_candidates.md",
            "ToDo候補",
            manual_todo_items,
        )

    if args.execute_commands:
        new_state, command_report = execute_at_commands(
            text=command_target_note,
            output_dir=output_dir,
            state=new_state,
            client=client,
            model=args.model,
            python_timeout=args.python_timeout,
            sync_google_todos=args.sync_google_todos,
            google_credentials=resolve_path(args.google_credentials),
            google_token=resolve_path(args.google_token),
            google_tasklist=args.google_tasklist,
            onenote_section=args.onenote_section,
            current_output_file=output_file,
            skip_ask=args.skip_ask,
            manual_google_tasks=args.manual_google_tasks,
            onenote_todo_items=explicit_todo_items,
        )
        should_update_onenote = command_report.get("onenote_status") == "not_requested"
        if should_update_onenote:
            onenote_status, onenote_detail = run_onenote_command(
                body="日誌要約の出力MarkdownでOneNoteページを更新",
                output_dir=output_dir,
                section_name=args.onenote_section,
                current_output_file=output_file,
                todo_items=explicit_todo_items,
            )
            command_report["onenote_status"] = onenote_status
            command_report["onenote_detail"] = onenote_detail

    save_state(state_dir, input_path, new_state)
    print(summary)
    print(f"\n保存しました: {output_file}")
    if manual_todo_path is not None:
        print(f"Google Tasks手動投入用: {manual_todo_path}")
    if args.execute_commands:
        print_command_report(command_report, output_dir)
        total_todos, missing_todos = audit_todo_registration(command_target_note, output_dir, manual_todo_path)
        print(f"@todo監査: 対象{total_todos}件 / 未確認{len(missing_todos)}件")
        for item in missing_todos:
            print(f"  - {item}")
        verification_output = build_command_report_text(
            output_file=output_file,
            command_report=command_report,
            total_todos=total_todos,
            missing_todos=missing_todos,
            manual_todo_path=manual_todo_path,
        )
        run_verification_agent(
            request=f"{args.date} 日誌生成・Google Tasks同期・OneNote更新",
            command=sys.argv,
            output=summary + "\n" + verification_output,
            exit_code=0,
        )
    else:
        run_verification_agent(
            request=f"{args.date} 日誌Markdown生成",
            command=sys.argv,
            output=summary + f"\n保存しました: {output_file}",
            exit_code=0,
        )
    return True


def watch_with_polling(args: argparse.Namespace, client: Anthropic) -> None:
    interval_seconds = max(1.0, args.interval_minutes * 60.0)
    print(f"[watch:poll] {args.input} を {args.interval_minutes} 分ごとに確認します。Ctrl+Cで終了。")
    try:
        while True:
            summarize_once(args, client, reason="polling")
            time.sleep(interval_seconds)
    except KeyboardInterrupt:
        print("\n監視を終了しました。")


def watch_with_watchdog(args: argparse.Namespace, client: Anthropic) -> None:
    try:
        from watchdog.observers import Observer
    except ImportError:
        print("[info] watchdog が見つかりません。pip install watchdog を実行してください。pollingに切り替えます。")
        watch_with_polling(args, client)
        return

    input_path = resolve_path(args.input)
    text_files = iter_text_files(input_path)
    watched_paths = {p.resolve() for p in text_files}
    watched_dirs = {p.parent.resolve() for p in text_files}

    def callback(reason: str) -> None:
        summarize_once(args, client, reason=reason)

    debounced = DebouncedEventHandler(
        watched_paths=watched_paths,
        callback=callback,
        debounce_seconds=args.debounce_seconds,
    )

    observer = Observer()
    for watched_dir in watched_dirs:
        observer.schedule(debounced.make_handler(), str(watched_dir), recursive=False)

    observer.start()
    print(f"[watch:event] 更新があった時だけ処理します。Ctrl+Cで終了。")
    for path in sorted(watched_paths):
        print(f"  - {path}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n監視を終了します。")
    finally:
        observer.stop()
        observer.join()


def print_rules() -> None:
    print("研究OS コマンド仕様 Version 1.0")
    print("通常文章: 日誌・要約・AI分析の対象")
    print("書式: @command 本文  または  ＠command の次行から本文")
    print("次の @ コマンドが出るまでを本文として扱います。")
    print()
    for name, description in SUPPORTED_AT_COMMANDS.items():
        status = "実行" if name in {"todo", "python", "ask", "命令", "onenote"} else "キュー保存"
        print(f"@{name:10s} {description} / {status}")
    print()
    print("例:")
    print("@todo STM測定継続")
    print("@python daily_paper_search.py")
    print("@ask")
    print("360Hzノイズの原因を考察して")
    print("＠命令")
    print("todo 試験問題を確認する")
    print("出力した.mdをOneNoteの日誌セクションに新しいノートとして追加して")
    print("ask STMの初期スキャン条件を整理して")
    print("@paper CO2RR graphene STM")


def main() -> None:
    load_agent_env()
    apply_secret_defaults()
    args = parse_args()
    if args.local_summary:
        args.skip_ask = True
    if args.show_rules:
        print_rules()
        return
    # commands-only/local-summary では通常Claude APIを使わない。
    # @askが現れた場合だけ run_ask_command 内で遅延作成する。
    client = None if (args.commands_only or args.local_summary) else llm_client.get_client()

    if args.watch:
        if not args.no_initial_run:
            summarize_once(args, client, reason="initial run")
        if args.poll:
            watch_with_polling(args, client)
        else:
            watch_with_watchdog(args, client)
    else:
        summarize_once(args, client)


if __name__ == "__main__":
    main()
