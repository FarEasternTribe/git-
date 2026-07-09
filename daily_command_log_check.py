from __future__ import annotations

import json
import os
import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


WORKSPACE_DIR = Path(__file__).resolve().parent
COMMAND_LOG_DIR = WORKSPACE_DIR / "agent_workspace" / "司令塔Agent" / "onenote_command_log"
SUMMARY_DIR = WORKSPACE_DIR / "agent_workspace" / "司令塔Agent" / "summaries"
SYNC_SCRIPT = WORKSPACE_DIR / "sync_onenote_command_log.ps1"
APPEND_SCRIPT = WORKSPACE_DIR / "append_onenote_command_log.ps1"
MIGRATION_CHECK = WORKSPACE_DIR / "migration_check.py"


def detect_device_label() -> str:
    configured = os.getenv("AGENT_DEVICE_LABEL", "").strip()
    if configured:
        return configured.strip("[]")
    computer_name = (os.getenv("COMPUTERNAME") or platform.node() or "").strip()
    upper_name = computer_name.upper()
    if "LENOVO" in upper_name:
        return "Lenovo"
    if "DESKTOP" in upper_name:
        return "Desktop"
    return computer_name or "UnknownPC"


def run_command(command: list[str], timeout: int = 180) -> tuple[int, str]:
    completed = subprocess.run(
        command,
        cwd=WORKSPACE_DIR,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=timeout,
    )
    return completed.returncode, completed.stdout.strip()


def load_state() -> dict:
    state_path = COMMAND_LOG_DIR / "state.json"
    if not state_path.exists():
        return {}
    return json.loads(state_path.read_text(encoding="utf-8-sig"))


def latest_page_record(state: dict) -> dict:
    pages = state.get("pages") or []
    if not pages:
        return {}
    return sorted(pages, key=lambda item: item.get("last_modified") or "", reverse=True)[0]


def load_changed_pages() -> list[dict]:
    path = COMMAND_LOG_DIR / "changed_pages.json"
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return []
    if isinstance(data, dict):
        return [data]
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    return []


def inspect_changed_pages(device: str, changed_pages: list[dict]) -> tuple[bool, str]:
    if not changed_pages:
        return True, "Changed pages: 0"

    other_device = "Lenovo" if device == "Desktop" else "Desktop" if device == "Lenovo" else ""
    summaries: list[str] = []
    other_device_hits = 0
    unlabeled = 0
    for record in changed_pages[:10]:
        title = str(record.get("title") or "")
        markdown = record.get("markdown")
        text = title
        if markdown and Path(markdown).exists():
            text += "\n" + Path(markdown).read_text(encoding="utf-8", errors="replace")
        has_desktop = "[Desktop]" in text or "- Device: Desktop" in text
        has_lenovo = "[Lenovo]" in text or "- Device: Lenovo" in text
        if other_device and ((other_device == "Lenovo" and has_lenovo) or (other_device == "Desktop" and has_desktop)):
            other_device_hits += 1
        if not has_desktop and not has_lenovo:
            unlabeled += 1
        summaries.append(f"- {title} / {record.get('last_modified')} / {markdown}")

    ok = unlabeled == 0
    detail_lines = [
        f"Changed pages: {len(changed_pages)}",
        f"Other-device changed pages detected: {other_device_hits}",
        f"Unlabeled changed pages: {unlabeled}",
        "",
        *summaries,
    ]
    return ok, "\n".join(detail_lines)


def inspect_latest_page(record: dict) -> tuple[bool, str]:
    markdown = record.get("markdown")
    if not markdown:
        return False, "Latest command-log page record has no markdown path."
    path = Path(markdown)
    if not path.exists():
        return False, f"Latest command-log markdown is missing: {path}"
    text = path.read_text(encoding="utf-8", errors="replace")
    known_device = any(label in text for label in ("[Lenovo]", "[Desktop]", "- Device: Lenovo", "- Device: Desktop"))
    if not known_device:
        return False, f"Latest command-log page has no known device label: {path}"
    return True, f"Latest command-log page has a device label: {record.get('title')}"


def write_summary(lines: list[str]) -> Path:
    SUMMARY_DIR.mkdir(parents=True, exist_ok=True)
    path = SUMMARY_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_daily_command_log_check.md"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def append_onenote_summary(device: str, ok: bool, summary_path: Path, actions: list[str], verification: list[str]) -> tuple[int, str]:
    summary = f"{device}日次検証: 命令したLog確認と再現性チェック {'OK' if ok else '要対応'}"
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
        "-NextSteps",
        "FAILがある場合は該当コードまたは認証・OneNote同期設定を修正して再実行する",
    ]
    return run_command(command, timeout=180)


def main() -> int:
    device = detect_device_label()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    checks: list[tuple[str, bool, str]] = []
    actions: list[str] = []

    sync_code, sync_output = run_command(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(SYNC_SCRIPT),
        ],
        timeout=240,
    )
    checks.append(("命令したLog同期", sync_code == 0, sync_output or f"exit={sync_code}"))
    actions.append("OneNote OpenAI_Agent1/命令したLogを同期")

    state = load_state()
    page_count = int(state.get("page_count") or 0)
    changed_count = int(state.get("changed_count") or 0)
    checks.append(("命令したLogページ数", page_count > 0, f"page_count={page_count}"))
    changed_pages = load_changed_pages()
    changed_ok, changed_detail = inspect_changed_pages(device, changed_pages)
    checks.append(("命令したLog差分確認", changed_ok, f"changed_count={changed_count}\n{changed_detail}"))
    record = latest_page_record(state)
    latest_ok, latest_detail = inspect_latest_page(record) if record else (False, "No page record found.")
    checks.append(("最新命令ログの端末ラベル", latest_ok, latest_detail))

    migration_code, migration_output = run_command(
        [sys.executable, str(MIGRATION_CHECK), "--deep", "--write-report"],
        timeout=240,
    )
    checks.append(("Lenovo/現端末での再現性チェック", migration_code == 0, migration_output or f"exit={migration_code}"))
    actions.append("migration_check.py --deep --write-report を実行")
    if changed_pages:
        actions.append("変更された命令したLogページを確認し、Desktop/Lenovo差分を監視サマリへ記録")

    ok = all(item[1] for item in checks)
    lines = [
        "# Daily Command Log Check",
        "",
        f"- Generated: {now}",
        f"- Device: {device}",
        f"- Workspace: `{WORKSPACE_DIR}`",
        f"- Result: {'OK' if ok else 'FAIL'}",
        "",
        "## Checks",
        "",
    ]
    for name, status, detail in checks:
        lines.extend([f"### {'[OK]' if status else '[FAIL]'} {name}", "", detail, ""])
    summary_path = write_summary(lines)

    verification = [f"{name}: {'OK' if status else 'FAIL'}" for name, status, _ in checks]
    append_code, append_output = append_onenote_summary(device, ok, summary_path, actions, verification)
    checks.append(("命令したLogへの日次検証記録", append_code == 0, append_output or f"exit={append_code}"))

    print("\n".join(lines))
    print(f"Summary written: {summary_path}")
    print("OneNote append:")
    print(append_output)
    return 0 if ok and append_code == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
