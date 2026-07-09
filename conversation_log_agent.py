from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from agent_file_logger import AGENT_LOG_ROOT, write_json_record, write_markdown_record
from agent_onenote_logger import write_agent_log


AGENT_NAME = "会話ログAgent"
WORKSPACE_DIR = Path(__file__).resolve().parent
CONVERSATION_JSONL = AGENT_LOG_ROOT / AGENT_NAME / "data" / "conversation_log.jsonl"


def append_jsonl(record: dict) -> Path:
    CONVERSATION_JSONL.parent.mkdir(parents=True, exist_ok=True)
    with CONVERSATION_JSONL.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
    return CONVERSATION_JSONL


def log_conversation(
    *,
    user_message: str,
    assistant_summary: str,
    actions: list[str],
    files: list[str],
    verification: list[str],
    next_steps: list[str],
    write_onenote: bool = False,
) -> tuple[Path, Path | None, str | None]:
    now = datetime.now()
    record = {
        "timestamp": now.isoformat(timespec="seconds"),
        "user_message": user_message,
        "assistant_summary": assistant_summary,
        "actions": actions,
        "files": files,
        "verification": verification,
        "next_steps": next_steps,
    }
    jsonl_path = append_jsonl(record)
    json_path = write_json_record(
        agent=AGENT_NAME,
        kind="conversation_record",
        request=user_message,
        data=record,
        created_at=now,
    )
    output = "\n".join(
        [
            "## Actions",
            *(f"- {item}" for item in actions),
            "",
            "## Files",
            *(f"- {item}" for item in files),
            "",
            "## Verification",
            *(f"- {item}" for item in verification),
            "",
            "## Next steps",
            *(f"- {item}" for item in next_steps),
            "",
            f"JSONL: {jsonl_path}",
            f"JSON: {json_path}",
        ]
    )
    md_path = write_markdown_record(
        agent=AGENT_NAME,
        kind="logs",
        request=user_message,
        decision_summary=assistant_summary,
        command=[],
        verification=verification,
        output=output,
        metadata=record,
        created_at=now,
    )

    onenote_detail = None
    if write_onenote:
        ok, detail = write_agent_log(
            agent=AGENT_NAME,
            request=user_message,
            decision_summary=assistant_summary,
            command=[],
            verification=verification,
            output=output,
            attach_pdfs=False,
        )
        onenote_detail = f"ok={ok} {detail}"

    return md_path, json_path, onenote_detail


def split_items(text: str | None) -> list[str]:
    if not text:
        return []
    return [item.strip() for item in text.split(";") if item.strip()]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="会話内容・決定事項・実行内容をAgent別Markdown/JSONLへ記録します。")
    parser.add_argument("--user-message", required=True)
    parser.add_argument("--assistant-summary", required=True)
    parser.add_argument("--actions", default="")
    parser.add_argument("--files", default="")
    parser.add_argument("--verification", default="")
    parser.add_argument("--next-steps", default="")
    parser.add_argument("--onenote", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    md_path, json_path, onenote_detail = log_conversation(
        user_message=args.user_message,
        assistant_summary=args.assistant_summary,
        actions=split_items(args.actions),
        files=split_items(args.files),
        verification=split_items(args.verification),
        next_steps=split_items(args.next_steps),
        write_onenote=args.onenote,
    )
    print(f"markdown={md_path}")
    print(f"json={json_path}")
    if onenote_detail:
        print(f"onenote={onenote_detail}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
