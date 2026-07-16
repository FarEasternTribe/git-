from __future__ import annotations

import argparse
import importlib.util
import json
import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from agent_config import env_file, google_credentials_path, google_token_path


WORKSPACE_DIR = Path(__file__).resolve().parent
DEFAULT_REPORT_DIR = WORKSPACE_DIR / "agent_workspace" / "司令塔Agent" / "summaries"
VENV_PYTHON = WORKSPACE_DIR / ".venv" / "Scripts" / "python.exe"
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"
BUNDLED_PYTHON = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)
BUNDLED_SITE_PACKAGES = BUNDLED_PYTHON.parent / "Lib" / "site-packages"


@dataclass
class Check:
    name: str
    status: str
    detail: str
    fix: str = ""


def status_icon(status: str) -> str:
    return {
        "OK": "[OK]",
        "WARN": "[WARN]",
        "FAIL": "[FAIL]",
        "INFO": "[INFO]",
    }.get(status, "[INFO]")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(WORKSPACE_DIR))
    except ValueError:
        return str(path)


def load_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def import_available(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def check_path_exists(name: str, path: Path, kind: str, fix: str = "") -> Check:
    if path.exists():
        if kind == "dir" and not path.is_dir():
            return Check(name, "FAIL", f"{rel(path)} exists but is not a directory.", fix)
        if kind == "file" and not path.is_file():
            return Check(name, "FAIL", f"{rel(path)} exists but is not a file.", fix)
        return Check(name, "OK", f"{rel(path)} exists.")
    return Check(name, "FAIL", f"{rel(path)} is missing.", fix)


def check_python() -> list[Check]:
    checks: list[Check] = []
    version = sys.version.split()[0]
    status = "OK" if sys.version_info >= (3, 9) else "FAIL"
    checks.append(
        Check(
            "Python version",
            status,
            f"{version} at {sys.executable}",
            "Install Python 3.9 or newer and rerun this script.",
        )
    )
    checks.append(Check("Operating system", "INFO", f"{platform.platform()}"))

    py_launcher = shutil.which("py")
    checks.append(
        Check(
            "Windows py launcher",
            "OK" if py_launcher else "WARN",
            py_launcher or "py launcher was not found in PATH.",
            "Install Python from python.org or use the full python.exe path.",
        )
    )
    checks.append(
        Check(
            "Project .venv python",
            "OK" if VENV_PYTHON.exists() else "WARN",
            str(VENV_PYTHON) if VENV_PYTHON.exists() else f"{rel(VENV_PYTHON)} was not found.",
            "Recreate the virtual environment on the new PC if project dependencies are missing.",
        )
    )
    checks.append(
        Check(
            "Codex bundled python",
            "OK" if BUNDLED_PYTHON.exists() else "INFO",
            str(BUNDLED_PYTHON) if BUNDLED_PYTHON.exists() else "Codex bundled Python was not found on this PC.",
            "This is optional, but useful inside the Codex desktop runtime.",
        )
    )
    return checks


def check_project_files() -> list[Check]:
    required_files = [
        "summarize_note5.py",
        "orchestrator_agent.py",
        "verification_agent.py",
        "conversation_log_agent.py",
        "agent_common.py",
        "agent_config.py",
        "agent_onenote_logger.py",
        "agent_file_logger.py",
        "agent_runtime.py",
        "llm_client.py",
        "daily_command_log_check.py",
        "mutual_command_log_monitor.py",
        "sync_google_todo_queue.py",
        "setup_google_tasks.py",
        "tools/add_synthesis_pdf_to_onenote.py",
        "tools/agent_common.ps1",
        "tools/organic_synthesis_agent.ps1",
        "tools/paper_search_agent.ps1",
        "tools/onenote_search_agent.ps1",
    ]
    required_dirs = [
        "agent_workspace",
        "rawtext",
        "日誌",
        "tools",
    ]
    checks: list[Check] = []
    for item in required_files:
        checks.append(
            check_path_exists(
                f"Required file: {item}",
                WORKSPACE_DIR / item,
                "file",
                "Copy the full OpenAI-Agent folder from OneDrive or restore the missing file.",
            )
        )
    for item in required_dirs:
        checks.append(
            check_path_exists(
                f"Required folder: {item}",
                WORKSPACE_DIR / item,
                "dir",
                "Copy the full OpenAI-Agent folder from OneDrive.",
            )
        )
    return checks


def check_env_and_auth() -> list[Check]:
    external_env = env_file()
    legacy_env = WORKSPACE_DIR / ".env"
    env_example = WORKSPACE_DIR / ".env.example"
    env = load_env_file(external_env)
    if not env:
        env = load_env_file(legacy_env)
    checks: list[Check] = [
        Check(
            "External agent.env",
            "OK",
            f"Configured: {external_env}" if external_env.exists() else "Not present; optional API integrations remain disabled.",
            "Create agent.env outside the workspace only when an explicit API integration is needed.",
        ),
        check_path_exists(".env.example", env_example, "file", "Restore .env.example from the project."),
    ]
    if legacy_env.exists():
        checks.append(Check("Legacy workspace .env", "WARN", "A secret-bearing .env remains inside the workspace.", "Move it to the external secret directory."))

    anthropic_key = env.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")
    checks.append(
        Check(
            "ANTHROPIC_API_KEY",
            "OK",
            "Configured for explicit API tools." if anthropic_key else "Not configured; journal generation remains local-only by default.",
            "Set ANTHROPIC_API_KEY only when an explicitly API-based tool (--api-summary etc.) is required.",
        )
    )

    google_sync = env.get("GOOGLE_TODO_SYNC", os.environ.get("GOOGLE_TODO_SYNC", "")).strip().lower()
    google_sync_enabled = google_sync in {"1", "true", "yes", "on"}
    checks.append(
        Check(
            "GOOGLE_TODO_SYNC",
            "OK",
            "Enabled." if google_sync_enabled else "Disabled; TODO items stay local/manual.",
            "Set GOOGLE_TODO_SYNC=1 only when Google Tasks synchronization is wanted.",
        )
    )

    def _resolve_from_env(key: str, default_path):
        raw = env.get(key) or os.environ.get(key)
        if raw:
            raw_path = Path(os.path.expandvars(raw)).expanduser()
            return raw_path if raw_path.is_absolute() else WORKSPACE_DIR / raw_path
        return default_path

    def _with_legacy_fallback(path: Path, legacy_name: str) -> Path:
        legacy = WORKSPACE_DIR / legacy_name
        return legacy if not path.exists() and legacy.exists() else path

    # 実行時は load_agent_env()/apply_secret_defaults() と同じ優先順
    # （env設定→外部secrets既定→レガシーworkspaceファイル）でパスを解決する。
    credentials = _with_legacy_fallback(_resolve_from_env("GOOGLE_TASKS_CREDENTIALS", google_credentials_path()), "credentials.json")
    token = _with_legacy_fallback(_resolve_from_env("GOOGLE_TASKS_TOKEN", google_token_path()), "token_google_tasks.json")
    if google_sync_enabled:
        checks.append(
            check_path_exists(
                "Google Tasks OAuth client",
                credentials,
                "file",
                "Copy credentials.json or download a new OAuth client JSON from Google Cloud.",
            )
        )
        checks.append(
            check_path_exists(
                "Google Tasks token",
                token,
                "file",
                "Run setup_google_tasks.py on the new PC to authorize Google Tasks.",
            )
        )
        checks.extend(check_google_json(credentials, token))
    else:
        checks.append(Check("Google Tasks credentials", "OK", "Not required while Google synchronization is disabled."))
    return checks


def check_google_json(credentials: Path, token: Path) -> list[Check]:
    checks: list[Check] = []
    if credentials.exists():
        try:
            data = json.loads(credentials.read_text(encoding="utf-8"))
            has_client = "installed" in data or "web" in data
            checks.append(
                Check(
                    "Google OAuth client JSON shape",
                    "OK" if has_client else "WARN",
                    "Looks like an OAuth client JSON." if has_client else "Missing installed/web key.",
                    "Use a Google OAuth client JSON, not a service-account JSON.",
                )
            )
        except Exception as exc:
            checks.append(Check("Google OAuth client JSON shape", "FAIL", f"Could not parse JSON: {exc}"))

    if token.exists():
        try:
            data = json.loads(token.read_text(encoding="utf-8"))
            expiry = data.get("expiry")
            refresh_token = bool(data.get("refresh_token"))
            if expiry:
                try:
                    expiry_dt = datetime.fromisoformat(expiry.replace("Z", "+00:00"))
                    now = datetime.now(timezone.utc)
                    status = "OK" if refresh_token or expiry_dt > now else "WARN"
                    detail = f"Expiry: {expiry}; refresh_token: {'yes' if refresh_token else 'no'}"
                except ValueError:
                    status = "WARN"
                    detail = f"Expiry value is not ISO format: {expiry}"
            else:
                status = "WARN"
                detail = f"refresh_token: {'yes' if refresh_token else 'no'}; expiry missing."
            checks.append(
                Check(
                    "Google Tasks token shape",
                    status,
                    detail,
                    "Run setup_google_tasks.py again if Google Tasks sync fails.",
                )
            )
        except Exception as exc:
            checks.append(Check("Google Tasks token shape", "FAIL", f"Could not parse JSON: {exc}"))
    return checks


def check_python_modules() -> list[Check]:
    modules = [
        ("openai", "OpenAI API client", "pip install openai"),
        ("googleapiclient", "Google API client", "pip install -r requirements-google-tasks.txt"),
        ("google_auth_oauthlib", "Google OAuth client", "pip install -r requirements-google-tasks.txt"),
        ("pypdf", "PDF text extraction", "pip install pypdf"),
    ]
    checks: list[Check] = []
    for module, label, fix in modules:
        available_here = import_available(module)
        available_elsewhere = module_available_in_project_runtime(module)
        status = "OK" if available_here or available_elsewhere else "WARN"
        if available_here:
            detail = f"{module} import available in current Python."
        elif available_elsewhere:
            detail = f"{module} import available in project .venv or Codex bundled Python."
        else:
            detail = f"{module} import not available in checked Python runtimes."
        checks.append(
            Check(
                f"Python module: {label}",
                status,
                detail,
                fix,
            )
        )
    return checks


def module_available_in_project_runtime(module_name: str) -> bool:
    if module_exists_in_site_packages(module_name):
        return True
    for python_path in (VENV_PYTHON, BUNDLED_PYTHON):
        if not python_path.exists() or str(python_path) == sys.executable:
            continue
        cmd = [
            str(python_path),
            "-c",
            f"import importlib.util; raise SystemExit(0 if importlib.util.find_spec({module_name!r}) else 1)",
        ]
        try:
            completed = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        except Exception:
            continue
        if completed.returncode == 0:
            return True
    return False


def module_exists_in_site_packages(module_name: str) -> bool:
    parts = module_name.split(".")
    candidates = [parts[0], parts[0].replace("_", "-")]
    for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
        if not site_packages.exists():
            continue
        for candidate in candidates:
            if (site_packages / candidate).exists():
                return True
            if any(site_packages.glob(f"{candidate}-*.dist-info")):
                return True
            if any(site_packages.glob(f"{candidate.replace('-', '_')}-*.dist-info")):
                return True
    return False


def check_onenote_com(deep: bool) -> list[Check]:
    checks: list[Check] = []
    if not import_available("win32com"):
        if not deep:
            return [
                Check(
                    "OneNote COM",
                    "INFO",
                    "Python win32com is not available. This workspace normally uses PowerShell COM; run with --deep to test that route.",
                    "Run migration_check.py --deep on Windows with OneNote Desktop installed.",
                )
            ]
        return check_powershell_onenote_com()

    if not deep:
        return [
            Check(
                "OneNote COM",
                "INFO",
                "win32com is available. Use --deep to instantiate OneNote.Application.",
                "Run python migration_check.py --deep on the new PC.",
            )
        ]

    try:
        import win32com.client  # type: ignore

        app = win32com.client.Dispatch("OneNote.Application")
        hierarchy = app.GetHierarchy("", 0)
        status = "OK" if hierarchy else "WARN"
        detail = "OneNote.Application responded." if hierarchy else "OneNote responded with empty hierarchy."
        checks.append(
            Check(
                "OneNote COM deep check",
                status,
                detail,
                "Install/sync OneNote desktop and open the target notebook once.",
            )
        )
    except Exception as exc:
        checks.append(
            Check(
                "OneNote COM deep check",
                "FAIL",
                f"Could not access OneNote.Application: {exc}",
                "Install OneNote desktop, sign in, sync notebooks, and rerun this script.",
            )
        )
    return checks


def check_powershell_onenote_com() -> list[Check]:
    powershell = shutil.which("powershell")
    if not powershell:
        return [
            Check(
                "OneNote PowerShell COM deep check",
                "FAIL",
                "powershell was not found in PATH.",
                "Run this workspace on Windows with PowerShell and OneNote Desktop installed.",
            )
        ]

    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "$one = New-Object -ComObject OneNote.Application; "
        "$xml = ''; "
        "$one.GetHierarchy('', 0, [ref]$xml); "
        "if ([string]::IsNullOrWhiteSpace($xml)) { exit 2 } "
        "Write-Output 'OneNote.Application responded via PowerShell COM.'",
    ]
    try:
        completed = subprocess.run(
            command,
            cwd=WORKSPACE_DIR,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=30,
        )
    except Exception as exc:
        return [
            Check(
                "OneNote PowerShell COM deep check",
                "FAIL",
                f"Could not run PowerShell COM check: {exc}",
                "Install OneNote desktop, sign in, sync notebooks, and rerun this script.",
            )
        ]

    output = completed.stdout.strip()
    return [
        Check(
            "OneNote PowerShell COM deep check",
            "OK" if completed.returncode == 0 else "FAIL",
            output or f"exit={completed.returncode}",
            "Install OneNote desktop, sign in, sync notebooks, and rerun this script.",
        )
    ]


def check_compile() -> list[Check]:
    files = [
        "summarize_note5.py",
        "orchestrator_agent.py",
        "verification_agent.py",
        "conversation_log_agent.py",
        "agent_onenote_logger.py",
        "daily_command_log_check.py",
        "mutual_command_log_monitor.py",
        "tools/add_synthesis_pdf_to_onenote.py",
    ]
    cmd = [sys.executable, "-m", "py_compile", *[str(WORKSPACE_DIR / item) for item in files]]
    try:
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    except Exception as exc:
        return [Check("Python syntax check", "FAIL", f"Could not run py_compile: {exc}")]
    if completed.returncode == 0:
        return [Check("Python syntax check", "OK", f"Compiled {len(files)} core scripts.")]
    detail = (completed.stderr or completed.stdout or "").strip()
    return [
        Check(
            "Python syntax check",
            "FAIL",
            detail[:1000] if detail else f"py_compile exited with {completed.returncode}.",
            "Fix the syntax error before migration.",
        )
    ]


def check_onedrive_location() -> list[Check]:
    text = str(WORKSPACE_DIR).casefold()
    if "onedrive" in text:
        return [Check("OneDrive location", "OK", str(WORKSPACE_DIR))]
    return [
        Check(
            "OneDrive location",
            "WARN",
            str(WORKSPACE_DIR),
            "For easy migration, place/sync this OpenAI-Agent folder under OneDrive.",
        )
    ]


def build_report(checks: list[Check]) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    counts = {status: sum(1 for check in checks if check.status == status) for status in ["OK", "WARN", "FAIL", "INFO"]}
    lines = [
        "# Codex Agent Migration Check",
        "",
        f"- Generated: {now}",
        f"- Workspace: `{WORKSPACE_DIR}`",
        f"- Python: `{sys.executable}`",
        f"- Summary: OK {counts['OK']} / WARN {counts['WARN']} / FAIL {counts['FAIL']} / INFO {counts['INFO']}",
        "",
        "## Results",
        "",
    ]
    for check in checks:
        lines.append(f"### {status_icon(check.status)} {check.name}")
        lines.append("")
        lines.append(f"- Status: {check.status}")
        lines.append(f"- Detail: {check.detail}")
        if check.fix:
            lines.append(f"- Fix: {check.fix}")
        lines.append("")
    return "\n".join(lines)


def print_report(checks: list[Check]) -> None:
    counts = {status: sum(1 for check in checks if check.status == status) for status in ["OK", "WARN", "FAIL", "INFO"]}
    print(
        f"Migration check: OK {counts['OK']} / WARN {counts['WARN']} / "
        f"FAIL {counts['FAIL']} / INFO {counts['INFO']}"
    )
    for check in checks:
        print(f"{status_icon(check.status)} {check.name}: {check.detail}")
        if check.fix and check.status in {"WARN", "FAIL"}:
            print(f"    fix: {check.fix}")


def write_report(report: str, report_dir: Path) -> Path:
    report_dir.mkdir(parents=True, exist_ok=True)
    path = report_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_migration_check.md"
    path.write_text(report, encoding="utf-8")
    return path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check whether this Codex agent workspace is ready to run on another PC.")
    parser.add_argument("--deep", action="store_true", help="Instantiate OneNote.Application through COM.")
    parser.add_argument("--write-report", action="store_true", help="Write a Markdown report under agent_workspace.")
    parser.add_argument("--report-dir", type=Path, default=DEFAULT_REPORT_DIR, help="Directory for --write-report output.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    checks: list[Check] = []
    checks.extend(check_python())
    checks.extend(check_onedrive_location())
    checks.extend(check_project_files())
    checks.extend(check_env_and_auth())
    checks.extend(check_python_modules())
    checks.extend(check_onenote_com(args.deep))
    checks.extend(check_compile())

    print_report(checks)
    if args.write_report:
        path = write_report(build_report(checks), args.report_dir)
        print(f"Report written: {path}")
    return 1 if any(check.status == "FAIL" for check in checks) else 0


if __name__ == "__main__":
    raise SystemExit(main())
