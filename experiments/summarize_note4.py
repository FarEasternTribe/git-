import argparse
import hashlib
import json
import os
import re
import shlex
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Iterable


WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)

for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from openai import OpenAI


DEFAULT_INPUT = "rawtext"
DEFAULT_OUTPUT_DIR = "日誌"
DEFAULT_STATE_DIR = ".note_agent_state"
DEFAULT_MODEL = "gpt-5"


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


COMMAND_BLOCK_PATTERNS = [
    re.compile(r"\[COMMAND\](.*?)\[/COMMAND\]", re.IGNORECASE | re.DOTALL),
    re.compile(r"\[命令\](.*?)\[/命令\]", re.DOTALL),
    re.compile(r"^T\s*$\n(.*?)\n^T\s*$", re.MULTILINE | re.DOTALL),
]

TODO_LINE_PATTERN = re.compile(
    r"^\s*(?:[-*・□☐✅]?\s*)?(?:TODO|ToDo|To do|To d 0|やること|タスク)[:：]?\s*(.*)$",
    re.IGNORECASE,
)


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
        help=f"OpenAI APIで使うモデル。省略時は {DEFAULT_MODEL}",
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
        "--sync-google-todos",
        action="store_true",
        default=os.getenv("GOOGLE_TODO_SYNC", "").lower() in {"1", "true", "yes", "on"},
        help="@todoをGoogle ToDo/Google Tasksへ即時登録する。失敗時はキューに保存",
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
    current_hash = hashlib.sha256(note.encode("utf-8")).hexdigest()

    # 追記型メモなら、前回長さ以降だけを処理する。
    if previous_length and len(note) >= previous_length:
        delta = note[previous_length:].strip()
    else:
        # ファイルが短くなった、または初回の場合は全体を処理。
        delta = note.strip()

    new_state = {
        **state,
        "processed_length": len(note),
        "full_hash": current_hash,
        "previous_hash": previous_hash,
        "last_processed_at": datetime.now().isoformat(timespec="seconds"),
    }
    return delta, new_state


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

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        match = TODO_LINE_PATTERN.match(stripped)
        if match:
            in_todo_section = True
            rest = match.group(1).strip()
            if rest:
                todos.append(rest)
            continue

        if in_todo_section:
            if stripped.startswith(("-", "*", "・", "□", "☐")):
                todos.append(stripped.lstrip("-*・□☐ ").strip())
            elif re.match(r"^[A-ZＡ-Ｚa-zａ-ｚ].{0,20}:$", stripped):
                in_todo_section = False

    # 「欠席学生にレポート課題送る」のような単独行も拾う
    for line in lines:
        stripped = line.strip(" ・-*□☐\t")
        if any(keyword in stripped for keyword in ("する", "送る", "記入", "修正", "実行", "チェック", "測定", "出席")):
            if 3 <= len(stripped) <= 80 and stripped not in todos:
                todos.append(stripped)

    return todos



AT_COMMAND_PATTERN = re.compile(r"^\s*[@＠]([^\s@＠]+)(?:\s+(.*))?$", re.IGNORECASE)


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
    lines = text.splitlines()
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
        while i < len(lines):
            next_line = lines[i]
            if AT_COMMAND_PATTERN.match(next_line):
                break
            block_lines.append(next_line)
            i += 1

        block = "\n".join(block_lines).strip()
        body = first_arg if first_arg else block
        if body:
            commands.append({"name": name, "body": body})

    return commands


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
            normalized_lines.append(line)
            continue

        raw_name = match.group(1).lower().strip()
        name = INSTRUCTION_COMMAND_ALIASES.get(raw_name, raw_name)
        rest = (match.group(2) or "").strip()
        normalized_lines.append(f"@{name} {rest}".rstrip())

    return "\n".join(normalized_lines).strip()


def command_hash(command: dict) -> str:
    raw = f"{command.get('name','')}\n{command.get('body','')}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


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
        item = line.strip().strip("-*・□☐✅ \t")
        if item:
            todos.append(item)
    if not todos and body.strip():
        todos.append(body.strip())
    return todos


def build_google_tasks_service(credentials_path: Path, token_path: Path):
    """Google Tasks API serviceを作る。依存ライブラリが無い場合は呼び出し側でログ化する。"""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/tasks"]
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path.exists():
                raise FileNotFoundError(
                    f"Google OAuthクライアントJSONが見つかりません: {credentials_path}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes)
            creds = flow.run_local_server(port=0)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("tasks", "v1", credentials=creds)


def add_google_task(
    title: str,
    output_dir: Path,
    credentials_path: Path,
    token_path: Path,
    tasklist_id: str,
) -> bool:
    try:
        service = build_google_tasks_service(credentials_path, token_path)
        created = (
            service.tasks()
            .insert(tasklist=tasklist_id, body={"title": title})
            .execute()
        )
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@todo Google Tasks登録",
            f"- {title}\n\nGoogle task id: `{created.get('id', '')}`",
        )
        return True
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
        return False


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


def run_ask_command(question: str, output_dir: Path, client: OpenAI, model: str) -> None:
    if client is None:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@ask 実行エラー",
            "OpenAI client is not available. Run without --commands-only or set OPENAI_API_KEY.",
        )
        return

    prompt = f"""
あなたは研究メモを解析するAIです。
以下の質問・メモに対して、事実、推測、次の行動を分けて日本語で詳しく答えてください。

質問・メモ:
{question}
"""
    try:
        response = client.responses.create(model=model, input=prompt)
        answer = response.output_text.strip()
        append_markdown_log(
            output_dir,
            "analysis.md",
            "@ask 解析結果",
            f"### 質問\n\n{question}\n\n### 回答\n\n{answer}",
        )
    except Exception as exc:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@ask 実行エラー",
            f"### 質問\n\n{question}\n\n### エラー\n\n{exc}",
        )




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
        "source": "summarize_note4.py",
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
    client: OpenAI,
    model: str,
    python_timeout: int,
    sync_google_todos: bool,
    google_credentials: Path,
    google_token: Path,
    google_tasklist: str,
    depth: int = 0,
) -> dict:
    if depth > 5:
        append_markdown_log(
            output_dir,
            "command_execution_log.md",
            "@命令 実行停止",
            "命令の入れ子が深すぎるため停止しました。",
        )
        return state

    commands = parse_at_commands(text)
    if not commands:
        return state

    executed_hashes = set(state.get("executed_command_hashes", []))
    new_hashes = []

    for command in commands:
        name = command["name"]
        body = command["body"].strip()
        h = command_hash(command)
        if h in executed_hashes:
            continue

        timestamp = datetime.now().isoformat(timespec="seconds")

        if name == "todo":
            for item in split_todo_body(body):
                synced = False
                if sync_google_todos:
                    synced = add_google_task(
                        title=item,
                        output_dir=output_dir,
                        credentials_path=google_credentials,
                        token_path=google_token,
                        tasklist_id=google_tasklist,
                    )
                record = {
                    "created_at": timestamp,
                    "title": item,
                    "source": "summarize_note.py @todo",
                    "status": "synced_google_tasks" if synced else "pending_google_todo_sync",
                }
                if not synced:
                    append_jsonl(output_dir, "google_todo_queue.jsonl", record)
                append_candidates(output_dir, "todo_candidates.md", "@todo 追加候補", [item])
            append_markdown_log(output_dir, "command_execution_log.md", "@todo 処理", body)

        elif name == "python":
            run_python_command(body, output_dir, timeout=python_timeout)

        elif name == "ask":
            run_ask_command(body, output_dir, client=client, model=model)

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
                state = execute_at_commands(
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
                    depth=depth + 1,
                )
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

    if new_hashes:
        state = {
            **state,
            "executed_command_hashes": sorted(executed_hashes.union(new_hashes))[-1000:],
            "last_command_processed_at": datetime.now().isoformat(timespec="seconds"),
        }
    return state


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


def summarize_once(args: argparse.Namespace, client: OpenAI, reason: str = "manual run") -> bool:
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
    if state.get("full_hash") == current_hash and args.delta_only:
        print(f"[skip] 更新なし: {datetime.now().strftime('%H:%M:%S')}")
        return False

    if args.delta_only:
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

    print(f"[run] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} / reason={reason}")

    if args.commands_only:
        new_state = execute_at_commands(
            text=target_note,
            output_dir=output_dir,
            state=new_state,
            client=client,
            model=args.model,
            python_timeout=args.python_timeout,
            sync_google_todos=args.sync_google_todos,
            google_credentials=resolve_path(args.google_credentials),
            google_token=resolve_path(args.google_token),
            google_tasklist=args.google_tasklist,
        )
        save_state(state_dir, input_path, new_state)
        print("[ok] コマンドだけ実行しました。")
        return True

    response = client.responses.create(
        model=args.model,
        input=build_prompt(target_note, args.date, delta_only=args.delta_only),
    )
    summary = response.output_text.strip() + "\n"

    output_file = write_summary(
        output_dir=output_dir,
        journal_date=args.date,
        summary=summary,
        overwrite=args.overwrite,
        append_output=args.append_output,
    )

    at_commands = parse_at_commands(target_note)

    if args.extract_commands:
        block_commands = extract_command_blocks(target_note)
        at_command_lines = [f"@{c['name']}\n{c['body']}" for c in at_commands]
        append_candidates(
            output_dir,
            "command_candidates.md",
            "命令候補",
            [*block_commands, *at_command_lines],
        )

    if args.extract_todos:
        todo_items = extract_todo_candidates(target_note)
        for command in at_commands:
            if command["name"] == "todo":
                todo_items.extend(split_todo_body(command["body"]))
        append_candidates(
            output_dir,
            "todo_candidates.md",
            "ToDo候補",
            todo_items,
        )

    if args.execute_commands:
        new_state = execute_at_commands(
            text=target_note,
            output_dir=output_dir,
            state=new_state,
            client=client,
            model=args.model,
            python_timeout=args.python_timeout,
            sync_google_todos=args.sync_google_todos,
            google_credentials=resolve_path(args.google_credentials),
            google_token=resolve_path(args.google_token),
            google_tasklist=args.google_tasklist,
        )

    save_state(state_dir, input_path, new_state)
    print(summary)
    print(f"\n保存しました: {output_file}")
    return True


def watch_with_polling(args: argparse.Namespace, client: OpenAI) -> None:
    interval_seconds = max(1.0, args.interval_minutes * 60.0)
    print(f"[watch:poll] {args.input} を {args.interval_minutes} 分ごとに確認します。Ctrl+Cで終了。")
    try:
        while True:
            summarize_once(args, client, reason="polling")
            time.sleep(interval_seconds)
    except KeyboardInterrupt:
        print("\n監視を終了しました。")


def watch_with_watchdog(args: argparse.Namespace, client: OpenAI) -> None:
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
        status = "実行" if name in {"todo", "python", "ask", "命令"} else "キュー保存"
        print(f"@{name:10s} {description} / {status}")
    print()
    print("例:")
    print("@todo STM測定継続")
    print("@python daily_paper_search.py")
    print("@ask")
    print("360Hzノイズの原因を考察して")
    print("＠命令")
    print("todo 試験問題を確認する")
    print("ask STMの初期スキャン条件を整理して")
    print("@paper CO2RR graphene STM")


def main() -> None:
    load_dotenv(WORKSPACE_DIR / ".env")
    args = parse_args()
    if args.show_rules:
        print_rules()
        return
    client = None if args.commands_only else OpenAI()

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
