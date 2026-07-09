from __future__ import annotations

import json
import re
import subprocess
from datetime import datetime
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parent
AGENT_LOG_ROOT = WORKSPACE_DIR / "agent_workspace"


def safe_name(text: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\s]+', "_", text).strip("._")
    return value or "agent"


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def agent_dir(agent: str) -> Path:
    path = AGENT_LOG_ROOT / safe_name(agent)
    (path / "logs").mkdir(parents=True, exist_ok=True)
    (path / "data").mkdir(parents=True, exist_ok=True)
    (path / "summaries").mkdir(parents=True, exist_ok=True)
    return path


def write_markdown_record(
    *,
    agent: str,
    kind: str,
    request: str,
    decision_summary: str = "",
    command: list[str] | None = None,
    verification: list[str] | None = None,
    output: str = "",
    metadata: dict | None = None,
    created_at: datetime | None = None,
) -> Path:
    created_at = created_at or datetime.now()
    base = agent_dir(agent)
    if kind not in {"logs", "data", "summaries"}:
        kind = "logs"
    folder = base / kind
    filename = f"{created_at.strftime('%Y%m%d_%H%M%S')}_{safe_name(request)[:80]}.md"
    path = folder / filename
    command = command or []
    verification = verification or []
    metadata = metadata or {}

    lines = [
        f"# {agent} {kind}",
        "",
        f"- 日付: {created_at.strftime('%Y-%m-%d')}",
        f"- 時刻: {created_at.strftime('%H:%M:%S')}",
        f"- Agent: {agent}",
        f"- 種別: {kind}",
        f"- 依頼: {request}",
        "",
        "## 判断概要",
        decision_summary or "記録なし",
        "",
        "## 実行コマンド",
        f"```powershell\n{subprocess.list2cmdline(command) if command else 'なし'}\n```",
        "",
        "## 検証項目",
    ]
    lines.extend(f"- {item}" for item in verification) if verification else lines.append("- なし")
    lines.extend([
        "",
        "## 実行出力・結果",
        "```text",
        output.strip() or "なし",
        "```",
    ])
    if metadata:
        lines.extend([
            "",
            "## Metadata",
            "```json",
            json.dumps(metadata, ensure_ascii=False, indent=2),
            "```",
        ])
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def write_json_record(
    *,
    agent: str,
    kind: str,
    request: str,
    data: dict,
    created_at: datetime | None = None,
) -> Path:
    created_at = created_at or datetime.now()
    base = agent_dir(agent)
    folder = base / "data"
    filename = f"{created_at.strftime('%Y%m%d_%H%M%S')}_{safe_name(request)[:80]}_{safe_name(kind)}.json"
    path = folder / filename
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path
