from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from conversation_log_agent import log_conversation
from agent_onenote_logger import write_agent_log
from verification_agent import verify_and_log


WORKSPACE_DIR = Path(__file__).resolve().parent
# サブプロセスは、agent.ps1 がこのorchestratorを起動したのと同じPython
# （通常はプロジェクト .venv）で実行する。sys.executable が取れない特殊
# ケースだけ py -3 にフォールバックする。
PYTHON_LAUNCHER = [sys.executable] if sys.executable else ["py", "-3"]

# 日誌・@コマンドの元メモを置くrawtextフォルダ。
# 既定は旧OpenAI_Agent側の共有rawtext。RAWTEXT_DIR環境変数で上書き可能。
DEFAULT_RAWTEXT_DIR = (
    r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)"
    r"\0000000000OpenAI_Agent\OpenAI-Agent\rawtext"
)
RAWTEXT_DIR = os.getenv("RAWTEXT_DIR", DEFAULT_RAWTEXT_DIR)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

JOURNAL_SCRIPT = WORKSPACE_DIR / "summarize_note5.py"
CLASSIFIER_SCRIPT = WORKSPACE_DIR / "指示作業Onenote分類.py"
CONVERSATION_LOG_SCRIPT = WORKSPACE_DIR / "conversation_log_agent.py"
EXPERIMENT_PPT_SCRIPT = WORKSPACE_DIR / "append_onenote_experiment_day_to_ppt.ps1"
EXPERIMENT_TRACKER_SCRIPT = WORKSPACE_DIR / "experiment_tracker.ps1"
GOOGLE_TODO_SYNC_SCRIPT = WORKSPACE_DIR / "sync_google_todo_queue.py"
DEVICE_MONITOR_SCRIPT = WORKSPACE_DIR / "mutual_command_log_monitor.py"
TOOLS_DIR = WORKSPACE_DIR / "tools"
PAPER_SEARCH_SCRIPT = TOOLS_DIR / "paper_search_agent.ps1"
PAPER_INDEX_SCRIPT = WORKSPACE_DIR / "paper_index_agent.py"
ROUTE_LOG = TOOLS_DIR / "orchestrator_agent_log.jsonl"


@dataclass
class Route:
    agent: str
    reason: str
    command: list[str]
    verification: list[str]
    requires_external_send: bool = False
    requires_write: bool = False
    clarification_question: str | None = None


@dataclass
class CouncilDecision:
    participants: list[str]
    status: str
    reasons: list[str]
    followup_command: list[str] | None = None


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text).casefold()


def clarification_route(text: str) -> Route:
    target = "添付ファイル" if ".pdf" in text.casefold() else "ご依頼"
    return Route(
        agent="確認Agent",
        reason="対象または完了条件を一意に判断できないため、実行前の確認が必要です。",
        command=[],
        verification=["回答を受け取るまで書き込みや外部送信を開始しない"],
        clarification_question=(
            f"{target}について、何をどこまで行いますか？ "
            "（例: 要約のみ／要約してINDEX登録／要約・INDEX登録・OneNote反映まで）"
        ),
    )


def extract_notebook(text: str) -> str | None:
    quoted = re.findall(r"[「『\"]([^」』\"]+)[」』\"]", text)
    if quoted:
        return quoted[0].strip()

    candidates = [
        "FarEasternTribe",
        "2026実験",
        "2026年実験",
        "2026年書き込みテスト",
        "2025年書込テスト",
        "2025年書き込みテスト",
        "2025年7月日誌",
        "2025年5月日誌",
        "2025年4月日誌",
        "OpenAI_agent1",
    ]
    normalized = normalize_text(text)
    for candidate in candidates:
        if normalize_text(candidate) in normalized:
            return candidate
    return None


def extract_search_query(text: str) -> str:
    patterns = [
        r"(.+?)を探",
        r"(.+?)について",
        r"(.+?)検索",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            query = match.group(1).strip(" 　。を")
            if query:
                return query
    return text.strip()


def extract_date_arg(text: str) -> str | None:
    # 年つきの完全な日付を最優先（2026-07-13 / 2026/07/13 / 2026年7月13日 / 20260713 等）。
    match = re.search(r"(20\d{2})[-/年.]?(\d{1,2})[-/月.]?(\d{1,2})", text)
    if match:
        year, month, day = match.groups()
        try:
            return datetime(int(year), int(month), int(day)).strftime("%Y-%m-%d")
        except ValueError:
            return None

    # 相対日付語（今日/本日=当日, 昨日=前日, 一昨日=2日前, 明日=翌日, 明後日=2日後）。
    # 「今日の実験ノートを転記して」「昨日の実験まとめて」等の自然文で対象日を解決する。
    # 「一昨日」は「昨日」を部分文字列に含むため、より長い語を先に判定する。
    relative_map = [
        (("今日", "本日", "きょう", "ほんじつ"), 0),
        (("一昨日", "おととい"), -2),
        (("昨日", "きのう", "前日"), -1),
        (("明後日", "あさって"), 2),
        (("明日", "あした", "あす"), 1),
    ]
    for words, delta in relative_map:
        if any(word in text for word in words):
            return (date.today() + timedelta(days=delta)).strftime("%Y-%m-%d")

    # 年なしの「月/日」指定（例: 7/13, 07/13, 7-13, 7月13日, 全角７／１３）は年を補う。
    # 「XX/YYの実験ノート転記して」のような自然文で対象日を指定できるようにする。
    # 小数(3.5g など)との誤認を避けるため、区切りに "." は使わない。
    normalized = text.translate(str.maketrans("０１２３４５６７８９／－", "0123456789/-"))
    md = re.search(r"(?<!\d)(\d{1,2})\s*[/月-]\s*(\d{1,2})日?(?!\d)", normalized)
    if md:
        month, day = int(md.group(1)), int(md.group(2))
        today = date.today()
        try:
            candidate = date(today.year, month, day)
        except ValueError:
            return None
        # 当年だと未来日になる場合は前年の実験ノートとみなす（過去メモの転記が主目的のため）。
        if candidate > today:
            try:
                candidate = date(today.year - 1, month, day)
            except ValueError:
                return None
        return candidate.strftime("%Y-%m-%d")

    return None


def extract_days_arg(text: str) -> int | None:
    # 「直近30日」「過去14日」「2週間」「1ヶ月」等から対象日数を推定する（進行中実験ボードの期間指定）。
    normalized = text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    m = re.search(r"(\d{1,3})\s*(日間|日|週間|週|ヶ月|カ月|か月|月)", normalized)
    if not m:
        return None
    n = int(m.group(1))
    unit = m.group(2)
    if n <= 0:
        return None
    if unit.startswith("日"):
        days = n
    elif "週" in unit:
        days = n * 7
    else:  # 月系
        days = n * 30
    return min(days, 365)


def journal_route(text: str, execute: bool) -> Route:
    python_command = PYTHON_LAUNCHER
    # ユーザー方針（2026-07-09決定）: 日誌は既定でローカル要約に固定する。
    # - Claude APIを呼ばない（--local-summary）＝課金なし
    # - Google TasksへAPI送信しない（--no-sync-google-todos --manual-google-tasks）
    #   ＝Googleへ送らず、コピペ用の手動投入用リストだけ生成
    # - @ask（Claude＋web検索）も無効（--skip-ask）
    # ※Claude要約は「api要約」「api-summary」等の明示指定があるときだけ
    #   --api-summary でオプトイン実行する（そのときだけ課金）。
    normalized = normalize_text(text)
    use_api_summary = any(word in normalized for word in ["api要約", "claudeapi使用", "api使用", "api-summary", "apisummary"])
    command = [*python_command, "-u", str(JOURNAL_SCRIPT), RAWTEXT_DIR]
    if use_api_summary:
        command.extend(["--api-summary", "--execute-commands", "--no-sync-google-todos", "--manual-google-tasks"])
    else:
        command.extend([
            "--local-summary", "--execute-commands",
            "--no-sync-google-todos", "--manual-google-tasks", "--skip-ask",
        ])
    if "上書き" in text or "overwrite" in text.casefold():
        command.append("--overwrite")
    if "差分" in text:
        command.append("--delta-only")
    if "強制" in text or "force" in text.casefold():
        command.append("--force")

    if use_api_summary:
        reason = "明示的なAPI要約指定による日誌生成です（Claude API使用・Google Tasks送信なし）。"
    else:
        reason = "ローカル要約による日誌生成です（Claude API・Google Tasks送信なし）。"
    return Route(
        agent="日誌Agent",
        reason=reason,
        command=command,
        verification=[
            "Markdown日誌が保存されたことを確認する",
            "末尾に生データセクションが含まれることを確認する",
            *([] if use_api_summary else ["外部APIを使わないローカル生成であることを確認する"]),
            "Google Tasks手動投入用TodoListが作成され、@todo全件が反映されたことを確認する",
            "OneNote本文に主要語（研究/生活など）が含まれることを確認する",
        ],
        requires_external_send=use_api_summary,
        requires_write=execute,
    )


def command_execution_route(text: str, execute: bool) -> Route:
    python_command = PYTHON_LAUNCHER
    # ユーザー方針: Google TasksへAPI送信しない（手動投入用リストのみ）。
    # @ask（Claude＋web検索）も無効。@todo/@python 等のローカル処理は従来通り。
    command = [
        *python_command,
        "-u",
        str(JOURNAL_SCRIPT),
        RAWTEXT_DIR,
        "--commands-only",
        "--execute-commands",
        "--no-sync-google-todos",
        "--manual-google-tasks",
        "--skip-ask",
    ]
    return Route(
        agent="日誌Agent",
        reason="rawtext内の@コマンド実行に関する依頼です（Google Tasks送信・@askなし）。",
        command=command,
        verification=[
            "@コマンド実行ログが出力されたことを確認する",
            "Google Tasks手動投入用TodoListが更新されたことを確認する",
            "エラーがないことを検証Agentで確認する",
        ],
        requires_external_send=False,
        requires_write=True,
    )


def google_tasks_sync_route(text: str) -> Route:
    python_command = PYTHON_LAUNCHER
    command = [*python_command, str(GOOGLE_TODO_SYNC_SCRIPT)]
    if "dry" in text.casefold() or "確認だけ" in text or "ドライラン" in text:
        command.append("--dry-run")
    return Route(
        agent="GoogleTasksAgent",
        reason="Google Tasksのキュー同期に関する依頼です。",
        command=command,
        verification=[
            "Google Tasksキュー同期コマンドが正常終了したことを確認する",
            "同期済み/未同期キューの状態を確認する",
        ],
        requires_write="--dry-run" not in command,
    )


def experiment_tracker_route(text: str) -> Route:
    # 進行中実験ボード（experiment_tracker.ps1）: OneNote実験ノートを段階推定して
    # 要対応順のボードをローカル出力する。読み取りのみ・課金なし。
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(EXPERIMENT_TRACKER_SCRIPT),
    ]
    # 「直近30日」「過去2週間」等の対象期間指定があれば -Days に渡す（無ければ既定21日）。
    days = extract_days_arg(text)
    if days:
        command += ["-Days", str(days)]
    return Route(
        agent="実験ノートAgent",
        reason="OneNoteの実験ノートから進行中実験ボードを生成する依頼です。",
        command=command,
        verification=[
            "OneNote FarEasternTribe/実験の直近ページを読めたことを確認する",
            "進行中実験ボード（Markdown）が生成されたことを確認する",
        ],
        requires_write=True,
    )


def experiment_ppt_route(text: str) -> Route:
    date_arg = extract_date_arg(text)
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(EXPERIMENT_PPT_SCRIPT),
    ]
    if date_arg:
        command += ["-Date", date_arg]
    if "force" in text.casefold() or "強制" in text or "再追記" in text:
        command.append("-Force")
    if "dry" in text.casefold() or "確認だけ" in text or "ドライラン" in text:
        command.append("-DryRun")
    return Route(
        agent="実験ノートAgent",
        reason="OneNoteの実験ノートをExperiment.pptxへ転記する依頼です。",
        command=command,
        verification=[
            "OneNote FarEasternTribe/実験の同日ページを読めたことを確認する",
            "Experiment.pptxへ日付・本文・画像スライドが追記されたことを確認する",
            "画像がある場合は抽出件数とスライド反映を確認する",
            "重複防止stateが更新されたことを確認する",
        ],
        requires_write="-DryRun" not in command,
    )


def device_monitor_route(text: str) -> Route:
    python_command = PYTHON_LAUNCHER
    command = [*python_command, str(DEVICE_MONITOR_SCRIPT)]
    return Route(
        agent="端末相互監視Agent",
        reason="Desktop/Lenovoの命令したLog同期、差分確認、再現性チェックに関する依頼です。",
        command=command,
        verification=[
            "OpenAI_Agent1/命令したLogを同期できたことを確認する",
            "Desktop/Lenovoの端末ラベル付きログを確認する",
            "migration_check.py --deep --write-report による再現性チェックを確認する",
            "監視結果を命令したLogへ記録したことを確認する",
        ],
        requires_write=True,
    )


def classify_route(text: str, execute: bool) -> Route:
    notebook = extract_notebook(text)
    command = [*PYTHON_LAUNCHER, str(CLASSIFIER_SCRIPT)]
    if notebook:
        command += ["--notebook", notebook]
    if execute:
        command.append("--execute")
    if "全件" in text or "作り直" in text or "full" in text.casefold():
        command.append("--full-rebuild")

    return Route(
        agent="OneNote分類Agent",
        reason="OneNoteノートブックの分類に関する依頼です。",
        command=command,
        verification=[
            "対象ページ数を確認する",
            "分類カード数が対象ページ数と一致することを確認する",
            "失敗件数が0であることを確認する",
        ],
        requires_write=execute,
    )


def search_route(text: str) -> Route:
    query = extract_search_query(text)
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(WORKSPACE_DIR / "tools" / "onenote_search_agent.ps1"),
        "-Query",
        query,
    ]
    return Route(
        agent="OneNote検索Agent",
        reason="OneNote内のページ検索に関する依頼です。",
        command=command,
        verification=[
            "検索結果の件数を確認する",
            "各結果にページ名、セクション、OneNoteリンクが含まれることを確認する",
        ],
    )


def paper_index_route(text: str) -> Route:
    # PaperIndexAgent はローカルのPDF索引・要約ツール（Claude API不使用・課金なし）。
    python_command = PYTHON_LAUNCHER
    command = [*python_command, "-u", str(PAPER_INDEX_SCRIPT)]

    search_match = re.search(r"(?:--search|search|検索)\s+[\"']([^\"']+)[\"']", text, flags=re.IGNORECASE)
    if not search_match:
        search_match = re.search(r"(?:--search|search|検索)\s+([^\s]+)", text, flags=re.IGNORECASE)
    if search_match:
        command.extend(["--search", search_match.group(1).strip()])
    elif any(word in normalize_text(text) for word in ["検索", "search"]):
        query = extract_search_query(text)
        if query and normalize_text(query) not in {"paperindex", "論文index", "論文索引"}:
            command.extend(["--search", query])

    if any(word in normalize_text(text) for word in ["初期化", "init", "作成", "更新", "再構築", "rebuild"]):
        command.append("--init")
    if "overwrite" in text.casefold() or "上書き" in text:
        command.append("--overwrite")

    if any(word in normalize_text(text) for word in ["要約", "summary", "summarize"]):
        command.append("--require-summary")

    def add_path_arg(raw_value: str) -> None:
        cleaned = raw_value.strip().strip("\"'")
        if not cleaned or cleaned.startswith("--"):
            return
        if cleaned.lower() in {"paper-index", "paperindex"}:
            return
        if search_match and cleaned == search_match.group(1).strip():
            return
        if cleaned not in command:
            command.append(cleaned)

    # Accept Windows/OneDrive paths that do not contain the word "paper".
    # PowerShell removes quotes before agent.ps1 receives argv, so the route
    # must recover full PDF paths from the joined request text.
    for match in re.finditer(r"[A-Za-z]:\\.*?\.pdf", text, flags=re.IGNORECASE):
        add_path_arg(match.group(0))
    for match in re.finditer(r"(?:\.\\|/)[^\r\n\"']*?\.pdf", text, flags=re.IGNORECASE):
        add_path_arg(match.group(0))

    quoted_paths = re.findall(r"[\"']([^\"']+\.pdf|[^\"']*papers?[^\"']*)[\"']", text, flags=re.IGNORECASE)
    for raw_path in quoted_paths:
        add_path_arg(raw_path)
    for raw_path in re.findall(r"(?<!\S)(?:\.\\|[A-Za-z]:\\|/)?[^\s\"']*(?:papers|paper)[^\s\"']*(?:\.pdf)?", text, flags=re.IGNORECASE):
        add_path_arg(raw_path)

    return Route(
        agent="PaperIndexAgent",
        reason="論文PDFライブラリの要約・Markdown索引・MASTER_INDEX/Excel/CSV更新に関する依頼です。",
        command=command,
        verification=[
            "papers/library/PDFs が存在することを確認する",
            "papers/library/index/MASTER_INDEX.md が存在し、件数が出力と整合することを確認する",
            "papers/library/INDEX.md、papers_index.xlsx、papers_index.csv が生成されることを確認する",
            "各PDFのMarkdown索引と summaries/ の要約ファイルを確認する",
            "検索モードでは該当IndexとDOI/PDF情報が返ることを確認する",
        ],
        requires_write="--search" not in command,
    )


def paper_search_route(text: str) -> Route:
    query = extract_search_query(text)
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(PAPER_SEARCH_SCRIPT),
        "-Query",
        query,
    ]
    return Route(
        agent="論文検索Agent",
        reason="論文・文献・根拠ソースの探索に関する依頼です。",
        command=command,
        verification=[
            "候補文献ごとにタイトル、著者、掲載誌、年を確認する",
            "ソースとなる文献のDOIを必ず添付する",
            "DOIが見つからない場合は、DOI未発見と探索範囲を明示する",
            "OneNoteリンクまたは外部文献リンクを含める",
        ],
    )


def organic_synthesis_route(text: str) -> Route:
    query = extract_search_query(text)
    if not query or normalize_text(query) in {"有機合成", "合成"}:
        query = "有機合成"
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(WORKSPACE_DIR / "tools" / "organic_synthesis_agent.ps1"),
        "-Query",
        query,
    ]
    return Route(
        agent="有機合成Agent",
        reason="合成ルート設計、必要文献探索、試薬・反応条件検討に関する依頼です。",
        command=command,
        verification=[
            "既存OneNoteから関連する合成メモ・論文メモを確認する",
            "本文だけでなくSupplementary Information/Supporting Information/ESI/SI/特許/PDF本文まで検索する",
            "ソースとなる文献のDOIを必ず添付する",
            "合成ルート候補、必要文献、試薬・条件検討ポイントを分けて出力する",
            "直接条件が見つかった場合は、試薬量、当量、溶媒、温度、時間、後処理、精製、収率、スペクトル、DOIを記録する",
            "OneNoteリンクが含まれることを確認する",
        ],
    )


def conversation_log_route(text: str, execute: bool) -> Route:
    command = [
        *PYTHON_LAUNCHER,
        str(CONVERSATION_LOG_SCRIPT),
        "--user-message",
        text,
        "--assistant-summary",
        "会話内容、決定事項、実行内容、検証結果、次の引き継ぎをMarkdown/JSONLに記録します。",
        "--actions",
        "会話ログをMarkdownに保存;会話ログをJSONLに追記;Agent別フォルダに保存",
        "--verification",
        "Markdownファイルが作成されること;JSONLに追記されること",
        "--next-steps",
        "重要な会話の節目でこのAgentを呼び出す",
    ]
    if execute or "onenote" in normalize_text(text):
        command.append("--onenote")
    return Route(
        agent="会話ログAgent",
        reason="会話内容や決定事項を継続ログとして保存する依頼です。",
        command=command,
        verification=[
            "Markdownログが作成されることを確認する",
            "JSONLログに追記されることを確認する",
            "必要に応じてOneNoteのAgent_会話ログAgentにも保存する",
        ],
        requires_write=True,
    )


def route_request(text: str, execute: bool) -> Route:
    normalized = normalize_text(text)
    if any(
        word in normalized
        for word in [
            "相互監視",
            "端末監視",
            "desktopとlenovo",
            "lenovoとdesktop",
            "lenovoでも同じ仕様",
            "device-monitor",
            "daily-command-log-check",
            "命令したlog差分",
        ]
    ):
        return device_monitor_route(text)
    if any(word in normalized for word in ["会話ログ", "ログ取", "リアルタイムログ", "会話を記録", "conversationlog"]):
        return conversation_log_route(text, execute)
    if any(word in normalized for word in ["@コマンド", "コマンド実行", "commands-only", "命令実行", "rawtext内の@"]):
        return command_execution_route(text, execute)
    if any(word in normalized for word in ["googletasks", "googletodo", "todo同期", "キュー同期"]):
        return google_tasks_sync_route(text)
    if any(
        word in normalized
        for word in [
            "進行中の実験",
            "進行中実験",
            "実験の進捗",
            "実験進捗",
            "実験ボード",
            "実験トラッカー",
            "実験の状況",
            "実験状況",
            "実験の棚卸",
            "どの反応",
        ]
    ):
        # 進行中実験ボードは、PPT転記(実験＋まとめ/転記)より先に判定する。
        # 「実験の進捗をまとめて」のように "まとめ" を含んでもボード生成を優先する。
        return experiment_tracker_route(text)
    if any(
        word in normalized
        for word in [
            "実験ppt",
            "experiment.pptx",
            "実験ノートagent",
            "実験ノートをppt",
            "実験ノートをpowerpoint",
            "実験ノート転記",
            "実験ppt作成",
            "experiment-onenote-day",
        ]
    ) or (
        # 「今日の実験まとめて」「実験をまとめて」等の自然文も OneNote実験ノート→PPT転記に回す。
        # 「実験」＋（まとめ／転記／PPT系）の同時出現を条件にして過検出を防ぐ
        # （日誌ルートの "要約" より前で判定される）。
        "実験" in normalized
        and any(
            word in normalized
            for word in ["まとめ", "転記", "ppt", "pptx", "パワポ", "パワーポイント", "スライド"]
        )
    ):
        return experiment_ppt_route(text)
    if ".pdf" in text.casefold() and any(word in text.casefold() for word in ["要約", "summary", "summarize"]):
        return paper_index_route(text)
    if any(word in normalized for word in ["paper-index", "paperindex", "master_index", "論文index", "論文索引", "pdf索引", "index化", "ｉｎｄｅｘ化", "論文pdfライブラリ", "論文ライブラリ"]):
        return paper_index_route(text)
    if any(word in normalized for word in ["論文検索", "文献検索", "論文を探", "文献を探", "papersearch", "literaturesearch"]):
        return paper_search_route(text)
    if any(
        word in normalized
        for word in [
            "有機合成",
            "合成法",
            "合成方法",
            "合成ルート",
            "反応条件",
            "試薬",
            "sonogashira",
            "鈴木",
            "薗頭",
            "negishi",
            "organic synthesis",
        ]
    ):
        return organic_synthesis_route(text)
    if any(word in normalized for word in ["分類", "classify"]):
        return classify_route(text, execute)
    if any(word in normalized for word in ["日誌", "要約", "生活ログ", "研究ログ", "summarize"]):
        return journal_route(text, execute)
    if any(word in normalized for word in ["探", "検索", "どこ", "onenoteから"]):
        return search_route(text)
    return clarification_route(text)


def append_route_log(request: str, route: Route, exit_code: int | None = None) -> None:
    TOOLS_DIR.mkdir(exist_ok=True)
    record = {
        "request": request,
        "agent": route.agent,
        "reason": route.reason,
        "command": route.command,
        "verification": route.verification,
        "requires_external_send": route.requires_external_send,
        "requires_write": route.requires_write,
        "clarification_question": route.clarification_question,
        "exit_code": exit_code,
    }
    with ROUTE_LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def print_route(route: Route) -> None:
    print(f"Agent: {route.agent}")
    print(f"Reason: {route.reason}")
    if route.requires_external_send:
        print("ExternalSend: yes")
    if route.requires_write:
        print("Write: yes")
    if route.command:
        print("Command:")
        print("  " + subprocess.list2cmdline(route.command))
    print("Verification:")
    for item in route.verification:
        print(f"  - {item}")
    if route.clarification_question:
        print(f"確認質問: {route.clarification_question}")


def run_command(command: list[str]) -> int:
    if not command:
        return 2
    completed = subprocess.run(command, cwd=WORKSPACE_DIR, text=True)
    return completed.returncode


def should_attach_pdfs(request: str, agent: str) -> bool:
    normalized = normalize_text(request)
    return agent == "有機合成Agent" or any(word in normalized for word in ["論文", "pdf", "paper"])


def requires_mandatory_verification(agent: str) -> bool:
    return agent in {"有機合成Agent", "論文検索Agent", "実験ノートAgent", "端末相互監視Agent"}


def should_auto_run(request: str, route: Route) -> bool:
    normalized = normalize_text(request)
    if route.agent == "日誌Agent" and "日誌" in normalized and "実行" in normalized:
        return True
    return False


def has_doi_text(output: str) -> bool:
    return re.search(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", output, flags=re.IGNORECASE) is not None


def build_agent_council_decision(
    *,
    request: str,
    route: Route,
    output: str,
    verification_status: str | None,
) -> CouncilDecision:
    participants = ["司令塔Agent", route.agent, "検証Agent"]
    if route.agent == "有機合成Agent" and "論文検索Agent" not in participants:
        participants.append("論文検索Agent")
    if route.agent == "論文検索Agent" and "有機合成Agent" not in participants:
        participants.append("有機合成Agent")

    reasons: list[str] = []
    followup_command: list[str] | None = None

    if verification_status is None:
        reasons.append("検証Agentの結果がないため、追加判断は保留です。")
        return CouncilDecision(participants, "needs_human", reasons)

    if verification_status != "ok":
        reasons.append(f"検証Agentステータスが {verification_status} のため、担当Agentの追加作業が必要です。")
        followup_command = build_repair_command(route, request, output)
        return CouncilDecision(participants, "needs_followup", reasons, followup_command)

    if route.agent in {"有機合成Agent", "論文検索Agent"}:
        if not has_doi_text(output):
            reasons.append("根拠DOIが出力に見つからないため、追加の文献・SI/PDF探索が必要です。")
            followup_command = build_repair_command(route, request, output)
            return CouncilDecision(participants, "needs_followup", reasons, followup_command)
        if any(term in output for term in ["DOI未発見", "未確認", "原文確認が必要"]):
            reasons.append("未確認事項が残っているため、SI/PDF/出版社ページまで追加確認する余地があります。")
            followup_command = build_repair_command(route, request, output)
            return CouncilDecision(participants, "needs_followup", reasons, followup_command)

    reasons.append("検証AgentがOKで、会議参加Agentから追加作業の必要は出ませんでした。")
    return CouncilDecision(participants, "complete", reasons)


def council_decision_text(decision: CouncilDecision) -> str:
    lines = [
        "Agent会議:",
        "参加Agent:",
        *(f"- {agent}" for agent in decision.participants),
        "",
        f"会議結論: {decision.status}",
        "理由:",
        *(f"- {reason}" for reason in decision.reasons),
    ]
    if decision.followup_command:
        lines.extend(["", "追加実行コマンド:", subprocess.list2cmdline(decision.followup_command)])
    return "\n".join(lines)


def write_council_log(request: str, route: Route, decision: CouncilDecision) -> None:
    ok, detail = write_agent_log(
        agent="司令塔Agent",
        request=f"Agent会議: {request}",
        decision_summary=f"{route.agent} の実行・検証後にAgent会議を行い、結論を {decision.status} と判定しました。",
        command=decision.followup_command or [],
        verification=[
            "担当Agentの実行結果を確認する",
            "検証Agentの判定を確認する",
            "追加作業の要否を各Agent観点で確認する",
        ],
        output=council_decision_text(decision),
        attach_pdfs=False,
    )
    status = "ok" if ok else "failed"
    print(f"CouncilLogOneNote: {status} {detail}")


def write_route_to_onenote(request: str, route: Route, output: str = "") -> None:
    orchestrator_ok, orchestrator_detail = write_agent_log(
        agent="司令塔Agent",
        request=request,
        decision_summary=f"{route.agent} に割り振りました。理由: {route.reason}",
        command=route.command,
        verification=[
            "依頼内容を解釈する",
            "担当Agentを選定する",
            "担当Agentの実行コマンドと検証項目を記録する",
        ],
        output=output or "司令塔が担当Agentを選定しました。",
        attach_pdfs=False,
    )
    orchestrator_status = "ok" if orchestrator_ok else "failed"
    print(f"OrchestratorLogOneNote: {orchestrator_status} {orchestrator_detail}")

    ok, detail = write_agent_log(
        agent=route.agent,
        request=request,
        decision_summary=route.reason,
        command=route.command,
        verification=route.verification,
        output=output or "司令塔が担当Agentを選定しました。",
        attach_pdfs=should_attach_pdfs(request, route.agent),
    )
    status = "ok" if ok else "failed"
    print(f"AgentLogOneNote: {status} {detail}")


def run_agent_command(command: list[str]) -> tuple[int, str]:
    if not command:
        return 2, ""
    completed = subprocess.run(
        command,
        cwd=WORKSPACE_DIR,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    output = "\n".join(part for part in [completed.stdout.strip(), completed.stderr.strip()] if part)
    return completed.returncode, output


def add_or_replace_arg(command: list[str], option: str, value: str) -> list[str]:
    repaired = list(command)
    if option in repaired:
        index = repaired.index(option)
        if index + 1 < len(repaired):
            repaired[index + 1] = value
        else:
            repaired.append(value)
    else:
        repaired.extend([option, value])
    return repaired


def build_repair_command(route: Route, request: str, output: str) -> list[str] | None:
    command = list(route.command)
    if not command:
        return None

    if route.agent == "OneNote分類Agent":
        if "--full-rebuild" not in command:
            command.append("--full-rebuild")
        return command

    if route.agent == "日誌Agent":
        if "--force" not in command:
            command.append("--force")
        return command

    if route.agent == "OneNote検索Agent":
        query = extract_search_query(request)
        broadened_query = f"{query} 関連 類義語 英語 日本語 OneNote"
        return add_or_replace_arg(command, "-Query", broadened_query)

    if route.agent == "論文検索Agent":
        query = extract_search_query(request)
        broadened_query = f"{query} DOI doi.org PMID arXiv journal title author source reference"
        return add_or_replace_arg(command, "-Query", broadened_query)

    if route.agent == "有機合成Agent":
        query = extract_search_query(request)
        broadened_query = (
            f"{query} DOI Supplementary Supporting Information SI ESI PDF patent "
            "試薬 溶媒 温度 時間 収率 NMR mmol 当量"
        )
        return add_or_replace_arg(command, "-Query", broadened_query)

    if route.agent == "会話ログAgent":
        if "--onenote" not in command and "OneNote" in output:
            command.append("--onenote")
        return command

    if route.agent == "実験ノートAgent":
        if "-DryRun" in command:
            command.remove("-DryRun")
        if "-Force" not in command:
            command.append("-Force")
        return command

    if route.agent == "端末相互監視Agent":
        # 端末相互監視Agentの警告（多くはOneNote COMが非対話セッションで
        # 使えないこと由来）は同一コマンドの再実行では直らないため、
        # 他の未対応Agentと同様に「修正コマンドを作れない」扱いにする。
        # これをせずに元コマンドをそのまま返すと、検証→警告→無意味な
        # 再実行→また警告、のループになる。
        return None

    return None


def write_repair_instruction_to_agent(request: str, route: Route, repair_instructions: list[str], repair_command: list[str]) -> None:
    output = "\n".join(
        [
            "検証Agentからの修正指示:",
            *(f"- {item}" for item in repair_instructions),
            "",
            "再実行コマンド:",
            subprocess.list2cmdline(repair_command),
        ]
    )
    ok, detail = write_agent_log(
        agent=route.agent,
        request=f"検証Agentからの修正指示: {request}",
        decision_summary="検証Agentが警告/失敗を検出したため、担当Agentに修正して再実行するよう指示しました。",
        command=repair_command,
        verification=["修正指示を確認する", "修正版コマンドを再実行する", "再実行後に検証Agentで再確認する"],
        output=output,
        attach_pdfs=False,
    )
    status = "ok" if ok else "failed"
    print(f"RepairInstructionLogOneNote: {status} {detail}")


def maybe_repair_and_reverify(
    *,
    request: str,
    route: Route,
    first_output: str,
    first_verification,
    max_attempts: int,
    write_onenote: bool,
) -> tuple[int | None, str, str | None]:
    if max_attempts <= 0 or first_verification.status == "ok":
        return None, "", None

    repair_command = build_repair_command(route, request, first_output)
    if not repair_command:
        print("AutoRepair: 修正版コマンドを作れないため、再実行は行いません。")
        return None, "", first_verification.status

    print("AutoRepair: 検証Agentの指示に基づき、担当Agentを修正条件で再実行します。")
    if write_onenote:
        write_repair_instruction_to_agent(
            request=request,
            route=route,
            repair_instructions=first_verification.repair_instructions,
            repair_command=repair_command,
        )
    print("RepairCommand:")
    print("  " + subprocess.list2cmdline(repair_command))

    repair_exit_code, repair_output = run_agent_command(repair_command)
    if repair_output:
        print(repair_output)
    print(f"RepairExitCode: {repair_exit_code}")

    repair_verification = verify_and_log(
        agent=route.agent,
        request=f"{request} [修正再実行]",
        command=repair_command,
        output=repair_output,
        exit_code=repair_exit_code,
    )
    print(f"VerificationAgentAfterRepair: {repair_verification.status}")
    return repair_exit_code, repair_output, repair_verification.status


def run_agent_council(
    *,
    request: str,
    route: Route,
    output: str,
    exit_code: int | None,
    verification_status: str | None,
    max_rounds: int,
    write_onenote: bool,
) -> tuple[int | None, str, str | None, str]:
    if max_rounds <= 0:
        return exit_code, output, verification_status, "skipped"

    decision = build_agent_council_decision(
        request=request,
        route=route,
        output=output,
        verification_status=verification_status,
    )
    print("\n--- Agent会議 ---")
    print(council_decision_text(decision))
    if write_onenote:
        write_council_log(request, route, decision)

    if decision.status != "needs_followup" or not decision.followup_command:
        return exit_code, output, verification_status, decision.status

    print("CouncilFollowup: 会議判断に基づき、追加作業を実行します。")
    print("CouncilFollowupCommand:")
    print("  " + subprocess.list2cmdline(decision.followup_command))
    followup_exit_code, followup_output = run_agent_command(decision.followup_command)
    if followup_output:
        print(followup_output)
    print(f"CouncilFollowupExitCode: {followup_exit_code}")

    followup_verification = verify_and_log(
        agent=route.agent,
        request=f"{request} [Agent会議後の追加作業]",
        command=decision.followup_command,
        output=followup_output,
        exit_code=followup_exit_code,
    )
    print(f"CouncilFollowupVerificationAgent: {followup_verification.status}")

    merged_output = "\n\n--- Agent会議後の追加作業 ---\n\n".join(
        part for part in [output, followup_output] if part
    )
    return followup_exit_code, merged_output, followup_verification.status, "followup_executed"


def write_automatic_conversation_log(
    *,
    request: str,
    route: Route,
    output: str,
    exit_code: int | None,
    verification_status: str | None,
    council_status: str | None = None,
) -> None:
    actions = [
        f"司令塔Agentが依頼を受け付けた: {request}",
        f"担当Agentを選定した: {route.agent}",
        f"選定理由を記録した: {route.reason}",
    ]
    if route.command:
        actions.append("担当Agentの実行コマンドを記録した")
    if exit_code is not None:
        actions.append(f"担当Agentの実行結果を記録した: exit_code={exit_code}")
    if verification_status:
        actions.append(f"検証Agentの結果を記録した: {verification_status}")
    if council_status:
        actions.append(f"Agent会議の結果を記録した: {council_status}")

    verification = list(route.verification)
    if verification_status:
        verification.append(f"検証Agentステータス: {verification_status}")
    if council_status:
        verification.append(f"Agent会議ステータス: {council_status}")
    if exit_code is not None:
        verification.append(f"終了コード: {exit_code}")

    files = [
        str(ROUTE_LOG),
        str(CONVERSATION_LOG_SCRIPT),
    ]

    md_path, json_path, _ = log_conversation(
        user_message=request,
        assistant_summary=(
            f"司令塔Agentが {route.agent} に割り振りました。"
            f"理由: {route.reason}"
        ),
        actions=actions,
        files=files,
        verification=verification,
        next_steps=["次回以降も司令塔Agent経由の依頼を自動で会話ログへ追記する"],
        write_onenote=False,
    )
    print(f"ConversationLog: markdown={md_path}")
    print(f"ConversationLog: json={json_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="自然文の依頼を、日誌・OneNote分類・OneNote検索などの専門Agentへ振り分ける司令塔Agent。"
    )
    parser.add_argument("request", nargs="*", help="依頼文。例: 2026実験を分類して")
    parser.add_argument("--execute", action="store_true", help="書き込み系の処理を実行する。未指定時は計画表示のみ")
    parser.add_argument("--run", action="store_true", help="選ばれたコマンドを実際に起動する")
    parser.add_argument("--no-onenote-log", action="store_true", help="AgentログをOneNoteに記録しない")
    parser.add_argument("--no-verify", action="store_true", help="実行後の検証Agentを呼ばない")
    parser.add_argument("--no-conversation-log", action="store_true", help="会話ログAgentへの自動記録を行わない")
    parser.add_argument("--no-auto-repair", action="store_true", help="検証警告/失敗時の自動修正再実行を行わない")
    parser.add_argument("--max-repair-attempts", type=int, default=1, help="検証警告/失敗時の修正再実行回数")
    parser.add_argument("--no-council", action="store_true", help="実行・検証後のAgent会議を行わない")
    parser.add_argument("--max-council-rounds", type=int, default=1, help="Agent会議後の追加作業実行回数")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    request = " ".join(args.request).strip()
    if not request:
        request = input("司令塔Agentへの依頼: ").strip()
    if not request:
        print("依頼文が空です。")
        return 2

    route = route_request(request, execute=args.execute)
    print_route(route)
    if route.clarification_question:
        append_route_log(request, route, 2)
        return 2
    auto_run = should_auto_run(request, route)
    if auto_run and not args.run:
        print("AutoRun: 日誌、実行ショートカットとして外部送信可で日誌Agentを実行します。")

    exit_code: int | None = None
    output = ""
    if args.run or auto_run:
        if route.command:
            exit_code, output = run_agent_command(route.command)
            if output:
                print(output)
        else:
            exit_code = 2
        print(f"ExitCode: {exit_code}")

    if not args.no_onenote_log:
        write_route_to_onenote(request, route, output)

    verification_status: str | None = None
    mandatory_verify = requires_mandatory_verification(route.agent)
    if args.no_verify and mandatory_verify:
        print(f"VerificationAgent: {route.agent} は根拠確認が必須のため --no-verify を無視します。")

    if (args.run or auto_run) and (not args.no_verify or mandatory_verify):
        verification = verify_and_log(
            agent=route.agent,
            request=request,
            command=route.command,
            output=output,
            exit_code=exit_code,
        )
        verification_status = verification.status
        print(f"VerificationAgent: {verification_status}")
        if not args.no_auto_repair and verification.status != "ok":
            repair_exit_code, repair_output, repair_verification_status = maybe_repair_and_reverify(
                request=request,
                route=route,
                first_output=output,
                first_verification=verification,
                max_attempts=args.max_repair_attempts,
                write_onenote=not args.no_onenote_log,
            )
            if repair_exit_code is not None:
                exit_code = repair_exit_code
                output = "\n\n--- 修正再実行 ---\n\n".join(part for part in [output, repair_output] if part)
                verification_status = repair_verification_status

    council_status: str | None = None
    if (args.run or auto_run) and verification_status is not None and not args.no_council:
        exit_code, output, verification_status, council_status = run_agent_council(
            request=request,
            route=route,
            output=output,
            exit_code=exit_code,
            verification_status=verification_status,
            max_rounds=args.max_council_rounds,
            write_onenote=not args.no_onenote_log,
        )

    append_route_log(request, route, exit_code)
    if not args.no_conversation_log:
        write_automatic_conversation_log(
            request=request,
            route=route,
            output=output,
            exit_code=exit_code,
            verification_status=verification_status,
            council_status=council_status,
        )
    return exit_code if exit_code is not None else 0


if __name__ == "__main__":
    raise SystemExit(main())
