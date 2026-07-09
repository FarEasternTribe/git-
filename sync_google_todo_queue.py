import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)

for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

from dotenv import load_dotenv

import summarize_note5 as summarize_note


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync pending @todo items from google_todo_queue.jsonl to Google Tasks."
    )
    parser.add_argument(
        "--queue",
        default="日誌/google_todo_queue.jsonl",
        help="Pending queue JSONL file.",
    )
    parser.add_argument(
        "--synced",
        default="日誌/google_todo_synced.jsonl",
        help="JSONL file where synced records are appended.",
    )
    parser.add_argument(
        "--credentials",
        default=os.getenv("GOOGLE_TASKS_CREDENTIALS", "credentials.json"),
        help="OAuth desktop client JSON downloaded from Google Cloud.",
    )
    parser.add_argument(
        "--token",
        default=os.getenv("GOOGLE_TASKS_TOKEN", "token_google_tasks.json"),
        help="OAuth token JSON.",
    )
    parser.add_argument(
        "--tasklist",
        default=os.getenv("GOOGLE_TASKS_LIST_ID", "@default"),
        help="Google Tasks tasklist id. Use @default for the default list.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be synced without adding tasks.",
    )
    return parser.parse_args()


def read_queue(path: Path) -> tuple[list[dict], list[str]]:
    if not path.exists():
        return [], []

    records = []
    invalid_lines = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            records.append(json.loads(line))
        except json.JSONDecodeError:
            invalid_lines.append(line)
    return records, invalid_lines


def append_jsonl(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def write_queue(path: Path, records: list[dict], invalid_lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        for line in invalid_lines:
            f.write(line + "\n")
        for record in records:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    tmp_path.replace(path)


def main() -> None:
    load_dotenv(summarize_note.WORKSPACE_DIR / ".env")
    args = parse_args()

    queue_path = summarize_note.resolve_path(args.queue)
    synced_path = summarize_note.resolve_path(args.synced)
    credentials_path = summarize_note.resolve_path(args.credentials)
    token_path = summarize_note.resolve_path(args.token)
    output_dir = queue_path.parent

    records, invalid_lines = read_queue(queue_path)
    if not records:
        print(f"[ok] No pending tasks in {queue_path}.")
        return

    if args.dry_run:
        print(f"[dry-run] {len(records)} pending task(s):")
        for record in records:
            print(f"- {record.get('title', '')}")
        return

    service = summarize_note.build_google_tasks_service(credentials_path, token_path)
    known_keys = summarize_note.load_known_google_task_keys(output_dir, args.tasklist, include_queue=False)
    remaining = []
    synced_count = 0
    skipped_count = 0

    for record in records:
        title = str(record.get("title", "")).strip()
        if not title:
            record["status"] = "failed_missing_title"
            remaining.append(record)
            continue

        dedupe_key = record.get("dedupe_key") or summarize_note.task_dedupe_key(title, args.tasklist)
        if dedupe_key in known_keys or summarize_note.google_task_title_exists(service, args.tasklist, title):
            skipped_record = {
                **record,
                "status": "skipped_duplicate_google_tasks",
                "skipped_at": datetime.now().isoformat(timespec="seconds"),
                "dedupe_key": dedupe_key,
            }
            append_jsonl(synced_path, skipped_record)
            summarize_note.append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@todo queued Google Tasks duplicate skip",
                f"- {title}",
            )
            skipped_count += 1
            continue

        try:
            created = (
                service.tasks()
                .insert(tasklist=args.tasklist, body={"title": title})
                .execute()
            )
            synced_record = {
                **record,
                "status": "synced_google_tasks",
                "synced_at": datetime.now().isoformat(timespec="seconds"),
                "google_task_id": created.get("id", ""),
                "dedupe_key": dedupe_key,
            }
            append_jsonl(synced_path, synced_record)
            known_keys.add(dedupe_key)
            summarize_note.append_markdown_log(
                output_dir,
                "command_execution_log.md",
                "@todo queued Google Tasks sync",
                f"- {title}\n\nGoogle task id: `{created.get('id', '')}`",
            )
            synced_count += 1
        except Exception as exc:
            record["status"] = "pending_google_todo_sync"
            record["last_error"] = str(exc)
            remaining.append(record)

    write_queue(queue_path, remaining, invalid_lines)
    print(f"[ok] synced={synced_count}, skipped_duplicate={skipped_count}, remaining={len(remaining)}")


if __name__ == "__main__":
    main()
