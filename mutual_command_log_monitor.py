from __future__ import annotations

import json
import re
import sys
from functools import partial
from datetime import datetime
from pathlib import Path

from agent_common import detect_device_label, run_workspace_command

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


WORKSPACE_DIR = Path(__file__).resolve().parent
COMMAND_LOG_DIR = WORKSPACE_DIR / "agent_workspace" / "司令塔Agent" / "onenote_command_log"
MONITOR_DIR = WORKSPACE_DIR / "agent_workspace" / "司令塔Agent" / "mutual_monitor"
SYNC_SCRIPT = WORKSPACE_DIR / "sync_onenote_command_log.ps1"
APPEND_SCRIPT = WORKSPACE_DIR / "append_onenote_command_log.ps1"
MIGRATION_CHECK = WORKSPACE_DIR / "migration_check.py"


def other_device_labels(device: str) -> list[str]:
    if device.casefold() == "lenovo":
        return ["Desktop"]
    if device.casefold() == "desktop":
        return ["Lenovo"]
    return ["Lenovo", "Desktop"]


run_command = partial(run_workspace_command, workspace=WORKSPACE_DIR, timeout=240)


def read_json(path: Path, default: dict) -> dict:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return default


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_command_log_state() -> dict:
    return read_json(COMMAND_LOG_DIR / "state.json", {})


def read_markdown(record: dict) -> str:
    markdown = record.get("markdown")
    if not markdown:
        return ""
    path = Path(markdown)
    if not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError as exc:
        return f"\n[ReadError: {exc.__class__.__name__}: {exc}]\n"


def page_device(record: dict, text: str) -> str:
    title = record.get("title") or ""
    for label in ("Lenovo", "Desktop"):
        if f"[{label}]" in title or f"[{label}]" in text or f"Device: {label}" in text:
            return label
    return "Unknown"


def extract_section_items(text: str, section_name: str) -> list[str]:
    pattern = re.compile(rf"^##\s+{re.escape(section_name)}\s*$", re.MULTILINE)
    match = pattern.search(text)
    if not match:
        return []
    rest = text[match.end() :]
    next_heading = re.search(r"^##\s+", rest, re.MULTILINE)
    block = rest[: next_heading.start()] if next_heading else rest
    items: list[str] = []
    for line in block.splitlines():
        stripped = line.strip()
        if stripped.startswith("- "):
            item = stripped[2:].strip()
            if item and item != "なし":
                items.append(item)
    return items


def resolve_file_reference(raw: str) -> tuple[str, Path | None, bool]:
    cleaned = raw.strip().strip("`").strip()
    cleaned = cleaned.split(" / ")[0].strip()
    if not cleaned or cleaned == "なし":
        return raw, None, True
    path = Path(cleaned)
    candidates: list[Path] = []
    if path.is_absolute():
        candidates.append(path)
    else:
        candidates.append(WORKSPACE_DIR / path)
        candidates.append(WORKSPACE_DIR / cleaned.replace("/", "\\"))
    for candidate in candidates:
        if candidate.exists():
            return raw, candidate, True
    return raw, candidates[0] if candidates else None, False


def append_onenote_result(
    device: str,
    ok: bool,
    summary_path: Path,
    actions: list[str],
    verification: list[str],
    required_on_other_device: list[str],
    next_steps: list[str],
) -> tuple[int, str]:
    summary = f"{device}相互監視: {'OK' if ok else '要対応'}"
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(APPEND_SCRIPT),
        "-Device",
        device,
        "-Summary",
        summary,
        "-Actions",
        "; ".join(actions),
        "-Files",
        str(summary_path),
        "-Verification",
        "; ".join(verification),
        "-RequiredOnOtherDevice",
        "; ".join(required_on_other_device) if required_on_other_device else "相手端末側で相互監視結果を確認する",
        "-NextSteps",
        "; ".join(next_steps) if next_steps else "なし",
    ]
    return run_command(command, timeout=180)


def main() -> int:
    device = detect_device_label()
    other_devices = set(other_device_labels(device))
    MONITOR_DIR.mkdir(parents=True, exist_ok=True)
    monitor_state_path = MONITOR_DIR / f"state_{device}.json"
    monitor_state = read_json(monitor_state_path, {"seen": {}})
    seen: dict[str, str] = monitor_state.setdefault("seen", {})

    sync_code, sync_output = run_command(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SYNC_SCRIPT)],
        timeout=240,
    )
    command_log_state = read_command_log_state()
    pages = command_log_state.get("pages") or []

    new_records: list[dict] = []
    inspected_records: list[dict] = []
    missing_files: list[str] = []
    existing_files: list[str] = []
    unreadable_pages: list[str] = []
    required_other_device: list[str] = []

    for record in pages:
        text = read_markdown(record)
        source_device = page_device(record, text)
        page_id = record.get("page_id") or record.get("title") or ""
        page_hash = record.get("hash") or ""
        if source_device not in other_devices:
            if page_id and page_hash:
                seen.setdefault(page_id, page_hash)
            continue
        if seen.get(page_id) == page_hash:
            continue
        new_records.append(record)
        read_error = text.strip() if text.lstrip().startswith("[ReadError:") else ""
        if read_error:
            unreadable_pages.append(f"{record.get('title')} -> {record.get('markdown')} ({read_error})")
        files = extract_section_items(text, "Files")
        file_results = []
        for item in files:
            raw, path, exists = resolve_file_reference(item)
            if path is None:
                continue
            result_line = f"{raw} -> {path}"
            if exists:
                existing_files.append(result_line)
            else:
                missing_files.append(result_line)
            file_results.append({"raw": raw, "path": str(path), "exists": exists})
        inspected_records.append(
            {
                "title": record.get("title"),
                "page_id": page_id,
                "hash": page_hash,
                "source_device": source_device,
                "files": file_results,
                "read_error": read_error,
                "actions": extract_section_items(text, "Actions"),
                "verification": extract_section_items(text, "Verification"),
                "required_on_other_device": extract_section_items(text, "Required On Other Device"),
            }
        )
        for item in extract_section_items(text, "Required On Other Device"):
            required_other_device.append(f"{source_device}: {record.get('title')} -> {item}")
        if page_id and page_hash:
            seen[page_id] = page_hash

    migration_code, migration_output = run_command(
        [sys.executable, str(MIGRATION_CHECK), "--deep", "--write-report"],
        timeout=240,
    )

    ok = sync_code == 0 and migration_code == 0 and not missing_files and not unreadable_pages
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary_lines = [
        "# Mutual Command Log Monitor",
        "",
        f"- Generated: {now}",
        f"- Device: {device}",
        f"- Watching: {', '.join(sorted(other_devices))}",
        f"- Result: {'OK' if ok else 'FAIL'}",
        f"- New or changed external pages: {len(new_records)}",
        "",
        "## Sync",
        "",
        f"- Exit: {sync_code}",
        "```text",
        sync_output,
        "```",
        "",
        "## New External Logs",
        "",
    ]
    if not inspected_records:
        summary_lines.append("- No new external logs")
    else:
        for item in inspected_records:
            summary_lines.extend(
                [
                    f"### {item['source_device']}: {item['title']}",
                    "",
                    f"- PageId: `{item['page_id']}`",
                    f"- Read: {'FAIL' if item['read_error'] else 'OK'}",
                    f"- File checks: {sum(1 for f in item['files'] if f['exists'])} OK / {sum(1 for f in item['files'] if not f['exists'])} missing",
                    f"- Required on this device: {len(item['required_on_other_device'])} item(s)",
                    "",
                ]
            )
            for required in item["required_on_other_device"]:
                summary_lines.append(f"  - {required}")
    summary_lines.extend(["", "## Missing Files", ""])
    summary_lines.extend(f"- {item}" for item in missing_files) if missing_files else summary_lines.append("- None")
    summary_lines.extend(["", "## Unreadable Pages", ""])
    summary_lines.extend(f"- {item}" for item in unreadable_pages) if unreadable_pages else summary_lines.append("- None")
    summary_lines.extend(["", "## Existing Files", ""])
    summary_lines.extend(f"- {item}" for item in existing_files[:50]) if existing_files else summary_lines.append("- None")
    summary_lines.extend(["", "## Required On This Device", ""])
    summary_lines.extend(f"- {item}" for item in required_other_device) if required_other_device else summary_lines.append("- None")
    summary_lines.extend(
        [
            "",
            "## Migration Check",
            "",
            f"- Exit: {migration_code}",
            "```text",
            migration_output,
            "```",
        ]
    )

    summary_path = MONITOR_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_mutual_monitor.md"
    summary_path.write_text("\n".join(summary_lines), encoding="utf-8")
    write_json(monitor_state_path, monitor_state)
    detail_path = MONITOR_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_mutual_monitor_records.json"
    write_json(detail_path, {"records": inspected_records, "missing_files": missing_files, "unreadable_pages": unreadable_pages})

    actions = [
        "命令したLogを同期",
        f"監視対象: {', '.join(sorted(other_devices))}",
        f"相手端末ログ {len(new_records)} 件を確認",
        "関連ファイルの存在確認",
        "migration_check.py --deep --write-report を実行",
    ]
    verification = [
        f"実行端末: {device}",
        f"監視対象: {', '.join(sorted(other_devices))}",
        f"OneNote同期: {'OK' if sync_code == 0 else 'FAIL'}",
        f"相手端末の新規/変更ログ: {len(new_records)} 件",
        f"関連ファイル不足: {len(missing_files)} 件",
        f"読取不可ページ: {len(unreadable_pages)} 件",
        f"相手端末からのRequired On Other Device: {len(required_other_device)} 件",
        f"再現性チェック: {'OK' if migration_code == 0 else 'FAIL'}",
    ]
    next_steps = []
    if missing_files:
        next_steps.append("不足ファイルをOneDrive同期または生成スクリプトで復元する")
    if unreadable_pages:
        next_steps.append("読取不可の同期ページをOneDriveでローカル保持にして再実行する")
    if migration_code != 0:
        next_steps.append("migration_checkのFAIL項目を修正して再実行する")
    if required_other_device:
        next_steps.append("Required On This Device欄の項目を確認し、必要なコード・runbook・設定を反映する")
    if not new_records:
        next_steps.append("新規の相手端末ログなし。監視継続")

    append_code, append_output = append_onenote_result(
        device,
        ok,
        summary_path,
        actions,
        verification,
        [
            "相手端末側で命令したLog差分を確認する",
            "Required On This Device欄に項目がある場合はコード・runbook・設定へ反映する",
        ],
        next_steps,
    )

    print("\n".join(summary_lines))
    print(f"Summary written: {summary_path}")
    print(f"Detail written: {detail_path}")
    print("OneNote append:")
    print(append_output)
    return 0 if ok and append_code == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
