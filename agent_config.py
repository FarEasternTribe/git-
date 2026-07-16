from __future__ import annotations

import os
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parent


def secret_dir() -> Path:
    configured = os.environ.get("OPENAI_AGENT_SECRET_DIR", "").strip()
    if configured:
        return Path(os.path.expandvars(configured)).expanduser().resolve()
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return (base / "OpenAI-Agent" / "secrets").resolve()


def env_file() -> Path:
    configured = os.environ.get("OPENAI_AGENT_ENV", "").strip()
    return Path(os.path.expandvars(configured)).expanduser().resolve() if configured else secret_dir() / "agent.env"


def google_credentials_path() -> Path:
    configured = os.environ.get("GOOGLE_TASKS_CREDENTIALS", "").strip()
    return Path(os.path.expandvars(configured)).expanduser().resolve() if configured else secret_dir() / "credentials.json"


def google_token_path() -> Path:
    configured = os.environ.get("GOOGLE_TASKS_TOKEN", "").strip()
    return Path(os.path.expandvars(configured)).expanduser().resolve() if configured else secret_dir() / "token_google_tasks.json"


def _load_plain_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def load_agent_env() -> Path | None:
    """Load external secrets first; support the legacy workspace .env during migration."""
    external = env_file()
    legacy = WORKSPACE_DIR / ".env"
    _load_plain_env(external)
    if legacy != external:
        _load_plain_env(legacy)
    return external if external.exists() else legacy if legacy.exists() else None


def apply_secret_defaults() -> None:
    # Migration safeguard: while the external secret directory is not populated yet,
    # keep using the legacy workspace files so existing Google Tasks auth keeps working.
    credentials = google_credentials_path()
    legacy_credentials = WORKSPACE_DIR / "credentials.json"
    if not credentials.exists() and legacy_credentials.exists():
        credentials = legacy_credentials
    token = google_token_path()
    legacy_token = WORKSPACE_DIR / "token_google_tasks.json"
    if not token.exists() and legacy_token.exists():
        token = legacy_token
    os.environ.setdefault("GOOGLE_TASKS_CREDENTIALS", str(credentials))
    os.environ.setdefault("GOOGLE_TASKS_TOKEN", str(token))
