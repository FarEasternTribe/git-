from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from agent_config import WORKSPACE_DIR, env_file, google_credentials_path, google_token_path, secret_dir


RUNTIME_DIR = WORKSPACE_DIR / ".agent_runtime"
STATE_FILE = RUNTIME_DIR / "state.json"
BACKUP_DIR = WORKSPACE_DIR / "backups"
MAX_RUNS = 300
BACKUP_ITEMS = [
    "agent.ps1", "agent_common.py", "agent_config.py", "agent_runtime.py",
    "orchestrator_agent.py", "verification_agent.py", "migration_check.py",
    "summarize_note5.py", "llm_client.py", "paper_index_agent.py",
    "setup_google_tasks.py", "sync_google_todo_queue.py", "CLAUDE_AGENT_RUNBOOK.md",
    "CLAUDE.md", "AGENT_OPERATION_POLICY.md", ".env.example", ".gitignore", "tests",
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_state() -> dict:
    if not STATE_FILE.exists():
        return {"schema_version": 1, "runs": []}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"schema_version": 1, "runs": [], "warning": "state file was unreadable"}


def save_state(state: dict) -> None:
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    state["updated_at"] = now_iso()
    state["runs"] = state.get("runs", [])[-MAX_RUNS:]
    fd, temp_name = tempfile.mkstemp(prefix="state-", suffix=".json", dir=RUNTIME_DIR)
    os.close(fd)
    temp = Path(temp_name)
    temp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, STATE_FILE)


def operation_key(action: str, arguments: str) -> str:
    normalized = " ".join(arguments.split()).casefold()
    return hashlib.sha256(f"{action.casefold()}\n{normalized}".encode("utf-8")).hexdigest()


def begin(action: str, arguments: str, dedupe_seconds: int, force: bool) -> int:
    state = load_state()
    key = operation_key(action, arguments)
    now = datetime.now(timezone.utc)
    if not force and dedupe_seconds > 0:
        for run in reversed(state.get("runs", [])):
            if run.get("operation_key") != key or run.get("status") != "succeeded":
                continue
            try:
                completed = datetime.fromisoformat(run["finished_at"])
            except (KeyError, ValueError):
                break
            if (now - completed).total_seconds() <= dedupe_seconds:
                print(json.dumps({"decision": "duplicate", "run_id": run["run_id"]}, ensure_ascii=False))
                return 10
            break
    run_id = uuid.uuid4().hex[:16]
    previous_attempts = [r for r in state.get("runs", []) if r.get("operation_key") == key]
    state.setdefault("runs", []).append({
        "run_id": run_id,
        "action": action,
        "arguments": arguments,
        "operation_key": key,
        "status": "running",
        "attempt": len(previous_attempts) + 1,
        "started_at": now.isoformat(),
    })
    save_state(state)
    print(json.dumps({"decision": "run", "run_id": run_id}, ensure_ascii=False))
    return 0


def finish(run_id: str, status: str, exit_code: int, message: str) -> int:
    state = load_state()
    for run in reversed(state.get("runs", [])):
        if run.get("run_id") == run_id:
            run.update(status=status, exit_code=exit_code, message=message, finished_at=now_iso())
            save_state(state)
            return 0
    print(f"run not found: {run_id}", file=sys.stderr)
    return 2


def print_status(as_json: bool = False) -> int:
    state = load_state()
    runs = [run for run in state.get("runs", []) if run.get("action") != "status"]
    latest_by_action: dict[str, dict] = {}
    for run in runs:
        latest_by_action[run.get("action", "unknown")] = run
    payload = {
        "updated_at": state.get("updated_at"),
        "counts": {name: sum(r.get("status") == name for r in runs) for name in ("running", "succeeded", "failed")},
        "latest_by_action": latest_by_action,
    }
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    print("Agent runtime status")
    print(f"runs: {len(runs)} / running: {payload['counts']['running']} / succeeded: {payload['counts']['succeeded']} / failed: {payload['counts']['failed']}")
    for action, run in sorted(latest_by_action.items()):
        stamp = run.get("finished_at") or run.get("started_at", "")
        print(f"- {action}: {run.get('status')} / attempt {run.get('attempt', 1)} ({stamp})")
    return 0


def iter_backup_files():
    for relative in BACKUP_ITEMS:
        path = WORKSPACE_DIR / relative
        if path.is_file():
            yield path
        elif path.is_dir():
            yield from (p for p in path.rglob("*") if p.is_file() and "__pycache__" not in p.parts)


def backup(output: str | None) -> int:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    destination = Path(output).resolve() if output else BACKUP_DIR / f"agent-config-{datetime.now():%Y%m%d-%H%M%S}.zip"
    files = sorted(set(iter_backup_files()))
    manifest = {"created_at": now_iso(), "workspace": str(WORKSPACE_DIR), "contains_secrets": False, "files": {}}
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            relative = path.relative_to(WORKSPACE_DIR).as_posix()
            data = path.read_bytes()
            archive.writestr(relative, data)
            manifest["files"][relative] = hashlib.sha256(data).hexdigest()
        archive.writestr("BACKUP_MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2))
    print(destination)
    return 0


def restore(archive_name: str, apply: bool, confirm: bool) -> int:
    archive_path = Path(archive_name).resolve()
    if not archive_path.exists():
        print(f"archive not found: {archive_path}", file=sys.stderr)
        return 2
    with zipfile.ZipFile(archive_path) as archive:
        names = [n for n in archive.namelist() if n != "BACKUP_MANIFEST.json"]
        unsafe = [n for n in names if Path(n).is_absolute() or ".." in Path(n).parts]
        if unsafe:
            print(f"unsafe archive members: {unsafe}", file=sys.stderr)
            return 3
        print(f"restore candidates: {len(names)}")
        for name in names:
            print(f"- {name}")
        if not apply:
            print("dry-run only; no files changed")
            return 0
        if not confirm:
            print("restore apply requires --confirm-restore", file=sys.stderr)
            return 4
        safety = BACKUP_DIR / f"pre-restore-{datetime.now():%Y%m%d-%H%M%S}"
        safety.mkdir(parents=True, exist_ok=True)
        for name in names:
            target = (WORKSPACE_DIR / name).resolve()
            if WORKSPACE_DIR not in target.parents:
                return 3
            if target.exists() and target.is_file():
                copy_target = safety / name
                copy_target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target, copy_target)
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(name) as source, target.open("wb") as dest:
                shutil.copyfileobj(source, dest)
        print(f"restored; previous files copied to {safety}")
    return 0


def secrets_status() -> int:
    paths = {"secret_dir": secret_dir(), "env": env_file(), "google_credentials": google_credentials_path(), "google_token": google_token_path()}
    print(f"secret directory: {paths['secret_dir']}")
    for name in ("env", "google_credentials", "google_token"):
        path = paths[name]
        location = "outside workspace" if WORKSPACE_DIR not in path.parents else "INSIDE WORKSPACE"
        print(f"- {name}: {'present' if path.exists() else 'missing'} / {location} / {path}")
    return 0 if all(WORKSPACE_DIR not in paths[name].parents for name in ("env", "google_credentials", "google_token")) else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Central run ledger, deduplication, backup, restore, and secret status.")
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("begin")
    p.add_argument("--action", required=True); p.add_argument("--arguments", default="")
    p.add_argument("--dedupe-seconds", type=int, default=0); p.add_argument("--force", action="store_true")
    p = sub.add_parser("finish")
    p.add_argument("--run-id", required=True); p.add_argument("--status", choices=["succeeded", "failed"], required=True)
    p.add_argument("--exit-code", type=int, default=0); p.add_argument("--message", default="")
    p = sub.add_parser("status"); p.add_argument("--json", action="store_true")
    p = sub.add_parser("backup"); p.add_argument("--output")
    p = sub.add_parser("restore"); p.add_argument("archive"); p.add_argument("--apply", action="store_true"); p.add_argument("--confirm-restore", action="store_true")
    sub.add_parser("secrets-status")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.command == "begin": return begin(args.action, args.arguments, args.dedupe_seconds, args.force)
    if args.command == "finish": return finish(args.run_id, args.status, args.exit_code, args.message)
    if args.command == "status": return print_status(args.json)
    if args.command == "backup": return backup(args.output)
    if args.command == "restore": return restore(args.archive, args.apply, args.confirm_restore)
    if args.command == "secrets-status": return secrets_status()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
