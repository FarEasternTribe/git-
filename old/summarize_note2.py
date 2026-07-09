import argparse
import hashlib
import json
import re
import sys
import time
from dataclasses import dataclass
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


@dataclass
class NoteSnapshot:
    text: str
    digest: str
    source_names: list[str]


@dataclass
class ParsedCommands:
    command_blocks: list[str]
    todo_items: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="テキストメモを日付付きの日誌Markdownに整理し、必要に応じて監視・差分処理・命令抽出を行います。"
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
        help="メモを監視し、更新があったときだけ処理する",
    )
    parser.add_argument(
        "--interval-minutes",
        type=float,
        default=30.0,
        help="--watch 時の確認間隔。省略時は30分",
    )
    parser.add_argument(
        "--once-if-changed",
        action="store_true",
        help="前回処理時から更新があった場合だけ1回処理して終了する",
    )
    parser.add_argument(
        "--delta-only",
        action="store_true",
        help="前回処理時との差分だけを要約対象にする。通常は全文を要約する",
    )
    parser.add_argument(
        "--state-dir",
        default=DEFAULT_STATE_DIR,
        help=f"更新検出用の状態ファイル保存先。省略時は {DEFAULT_STATE_DIR}",
    )
    parser.add_argument(
        "--execute-safe-commands",
        action="store_true",
        help="安全な内蔵命令だけ実行する。未指定なら命令候補として記録のみ",
    )
    return parser.parse_args()


def build_prompt(note: str, journal_date: str, parsed: ParsedCommands | None = None) -> str:
    command_text = ""
    if parsed:
        if parsed.todo_items:
            command_text += "\n\n抽出されたTODO候補:\n" + "\n".join(f"- {x}" for x in parsed.todo_items)
        if parsed.command_blocks:
            command_text += "\n\n抽出された命令ブロック候補:\n" + "\n".join(
                f"---\n{x}" for x in parsed.command_blocks
            )

    return f"""
あなたは、研究者の日誌・研究メモ・生活ログを整理する専属AI秘書です。
以下の一日のテキストメモを、後から読み返しやすい日本語の日誌Markdownに整理してください。

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
- TODOと命令は、通常の日誌とは別に「実行管理」の観点でも整理してください。
- 危険な命令、曖昧な命令、外部送信を伴う命令は、自動実行せず「要確認」として扱ってください。

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

## 8. 明日以降のTODO
- 研究
- 生活
- 連絡・事務
- 体調管理

## 9. 命令・自動化候補
- メモ中の命令ブロック
- 自動実行してよい可能性があるもの
- 手動確認が必要なもの
- Google Todo / OneNote / 解析コードなどへの連携候補

## 10. 確認したいこと
意味が曖昧な記述、追加で確認すべき点を列挙してください。

## 11. ChatGPTによる詳細分析
以下を含めて、できるだけ詳しく分析してください。

### 11.1 今日の全体像
### 11.2 研究面の分析
### 11.3 生活・体調面の分析
### 11.4 心理面の分析
### 11.5 対人関係・ストレス構造の分析
### 11.6 明日に向けた行動戦略
### 11.7 注意すべきリスク
### 11.8 良かった点
### 11.9 次の一手

一日のテキストメモ:
{note}
{command_text}
"""


def resolve_path(path_like: str) -> Path:
    path = Path(path_like)
    if not path.is_absolute():
        path = WORKSPACE_DIR / path
    return path


def iter_text_files(input_path: Path) -> Iterable[Path]:
    if input_path.is_file():
        yield input_path
        return
    if not input_path.is_dir():
        raise FileNotFoundError(f"入力ファイルまたはフォルダが見つかりません: {input_path}")
    text_files = sorted(input_path.glob("*.txt"))
    if not text_files:
        raise FileNotFoundError(f"フォルダ内に .txt ファイルがありません: {input_path}")
    yield from text_files


def read_note_snapshot(input_path: Path) -> NoteSnapshot:
    chunks: list[str] = []
    source_names: list[str] = []
    for text_file in iter_text_files(input_path):
        content = text_file.read_text(encoding="utf-8-sig").strip()
        if not content:
            continue
        source_names.append(text_file.name)
        if input_path.is_file():
            chunks.append(content)
        else:
            chunks.append(f"--- {text_file.name} ---\n{content}")

    if not chunks:
        raise ValueError(f"入力メモが空です: {input_path}")

    text = "\n\n".join(chunks)
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
    return NoteSnapshot(text=text, digest=digest, source_names=source_names)


def state_file_path(state_dir: Path, input_path: Path) -> Path:
    key = hashlib.sha256(str(input_path.resolve()).encode("utf-8")).hexdigest()[:16]
    return state_dir / f"{key}.json"


def load_state(state_file: Path) -> dict:
    if not state_file.exists():
        return {}
    try:
        return json.loads(state_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_state(state_file: Path, snapshot: NoteSnapshot, processed_text: str) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    state = {
        "digest": snapshot.digest,
        "processed_at": datetime.now().isoformat(timespec="seconds"),
        "source_names": snapshot.source_names,
        "processed_text": processed_text,
    }
    state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def extract_delta(current_text: str, previous_text: str | None) -> str:
    if not previous_text:
        return current_text
    if current_text.startswith(previous_text):
        return current_text[len(previous_text):].strip()
    # 途中編集された場合は、差分推定で誤るより全文を再処理する。
    return current_text


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


def append_section(path: Path, title: str, lines: list[str]) -> None:
    if not lines:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    body = [f"\n## {timestamp} {title}", *lines, ""]
    with path.open("a", encoding="utf-8-sig") as f:
        f.write("\n".join(body))


def parse_command_blocks(note: str) -> list[str]:
    patterns = [
        r"\[COMMAND\](.*?)\[/COMMAND\]",
        r"\[命令\](.*?)\[/命令\]",
        r"```command\s*(.*?)```",
        r"```命令\s*(.*?)```",
    ]
    blocks: list[str] = []
    for pattern in patterns:
        blocks.extend(x.strip() for x in re.findall(pattern, note, flags=re.DOTALL | re.IGNORECASE) if x.strip())

    # 音声入力で「T」で囲む運用を想定。ただし誤検出しやすいので、単独行Tのみを対象にする。
    blocks.extend(
        x.strip()
        for x in re.findall(r"(?m)^T\s*$\n(.*?)\n^T\s*$", note, flags=re.DOTALL)
        if x.strip()
    )
    return blocks


def extract_todo_items(note: str) -> list[str]:
    todo_items: list[str] = []
    in_todo_area = False
    for raw_line in note.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if re.search(r"to\s*do|todo|やること", line, flags=re.IGNORECASE):
            in_todo_area = True
            # 見出し行自体はTODOにしない。
            if re.fullmatch(r"[\w\s:：.。-]*(to\s*do|todo|やること)[\w\s:：.。-]*", line, flags=re.IGNORECASE):
                continue
        if in_todo_area and re.match(r"^(#|##|\[|【)", line):
            in_todo_area = False
        if in_todo_area:
            cleaned = re.sub(r"^[-・*\d.、\s]+", "", line).strip()
            if cleaned and len(cleaned) >= 2:
                todo_items.append(cleaned)

    # 重複削除しつつ順序保持
    seen: set[str] = set()
    unique: list[str] = []
    for item in todo_items:
        if item not in seen:
            seen.add(item)
            unique.append(item)
    return unique


def parse_commands(note: str) -> ParsedCommands:
    return ParsedCommands(
        command_blocks=parse_command_blocks(note),
        todo_items=extract_todo_items(note),
    )


def record_or_execute_commands(
    parsed: ParsedCommands,
    output_dir: Path,
    execute_safe_commands: bool,
) -> None:
    if parsed.todo_items:
        todo_lines = [f"- [ ] {item}" for item in parsed.todo_items]
        append_section(output_dir / "todo_candidates.md", "TODO候補", todo_lines)

    if parsed.command_blocks:
        command_lines: list[str] = []
        for i, block in enumerate(parsed.command_blocks, start=1):
            command_lines.append(f"### 命令候補 {i}")
            command_lines.append(block)
            command_lines.append("")
        append_section(output_dir / "command_candidates.md", "命令候補", command_lines)

    if execute_safe_commands and parsed.todo_items:
        # 現段階の安全な自動実行は、ローカルのGoogle Todo投入候補ファイルを作るところまで。
        # 実際のGoogle Todo API登録、メール送信、ファイル削除などは手動確認後に別スクリプトで行う。
        google_todo_lines = [json.dumps({"title": item}, ensure_ascii=False) for item in parsed.todo_items]
        append_section(output_dir / "google_todo_import_candidates.jsonl", "Google Todo投入候補", google_todo_lines)


def summarize_note(client: OpenAI, model: str, note: str, journal_date: str, parsed: ParsedCommands) -> str:
    response = client.responses.create(
        model=model,
        input=build_prompt(note, journal_date, parsed),
    )
    return response.output_text.strip() + "\n"


def process_once(args: argparse.Namespace, *, force: bool = False) -> bool:
    load_dotenv(WORKSPACE_DIR / ".env")

    input_path = resolve_path(args.input)
    output_dir = resolve_path(args.output_dir)
    state_dir = resolve_path(args.state_dir)
    output_dir.mkdir(exist_ok=True)

    snapshot = read_note_snapshot(input_path)
    state_file = state_file_path(state_dir, input_path)
    state = load_state(state_file)

    if not force and state.get("digest") == snapshot.digest:
        print(f"更新なし: {input_path}")
        return False

    note_for_summary = snapshot.text
    if args.delta_only:
        note_for_summary = extract_delta(snapshot.text, state.get("processed_text"))
        if not note_for_summary.strip():
            print(f"差分なし: {input_path}")
            save_state(state_file, snapshot, snapshot.text)
            return False

    parsed = parse_commands(note_for_summary)
    record_or_execute_commands(parsed, output_dir, args.execute_safe_commands)

    client = OpenAI()
    summary = summarize_note(client, args.model, note_for_summary, args.date, parsed)

    output_file = unique_output_path(output_dir, args.date, args.overwrite)
    output_file.write_text(summary, encoding="utf-8-sig")

    save_state(state_file, snapshot, snapshot.text)

    print(summary)
    print(f"\n保存しました: {output_file}")
    if parsed.todo_items:
        print(f"TODO候補を保存しました: {output_dir / 'todo_candidates.md'}")
    if parsed.command_blocks:
        print(f"命令候補を保存しました: {output_dir / 'command_candidates.md'}")
    return True


def watch_loop(args: argparse.Namespace) -> None:
    interval_seconds = max(60, int(args.interval_minutes * 60))
    print(f"監視開始: {resolve_path(args.input)} / 間隔 {interval_seconds} 秒")
    while True:
        try:
            process_once(args)
        except Exception as exc:
            print(f"エラー: {exc}", file=sys.stderr)
        time.sleep(interval_seconds)


def main() -> None:
    args = parse_args()
    if args.watch:
        watch_loop(args)
    elif args.once_if_changed:
        process_once(args)
    else:
        process_once(args, force=True)


if __name__ == "__main__":
    main()
