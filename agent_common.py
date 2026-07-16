from __future__ import annotations

import os
import platform
import re
import subprocess
import tempfile
from pathlib import Path


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


def run_workspace_command(command: list[str], workspace: Path, timeout: int) -> tuple[int, str]:
    completed = subprocess.run(
        command,
        cwd=workspace,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=timeout,
    )
    return completed.returncode, completed.stdout.strip()


def ps_single_quoted(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def escape_cdata(text: str) -> str:
    return text.replace("]]>", "]]]]><![CDATA[>")


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def run_powershell_script(script: str, prefix: str = "onenote_classify_") -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory(prefix=prefix) as tmp:
        path = Path(tmp) / "script.ps1"
        path.write_text(script, encoding="utf-8-sig")
        return subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(path)],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
