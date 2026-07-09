from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path

from agent_file_logger import write_json_record, write_markdown_record


WORKSPACE_DIR = Path(__file__).resolve().parent
ROUTE_LOG = WORKSPACE_DIR / "tools" / "orchestrator_agent_log.jsonl"


MANUAL_BACKFILL = [
    {
        "agent": "有機合成Agent",
        "kind": "summaries",
        "request": "2-buthoxynaphthalene / 2-butoxynaphthalene の合成法調査",
        "decision_summary": (
            "OneNote内に直接一致は見つからず、外部検索で Nature Communications "
            "10.1038/s41467-024-50086-6 のSupplementary Informationに "
            "2-Butoxynaphthalene の直接条件があることを確認した。"
        ),
        "verification": [
            "本文だけではなくSupplementary Informationまで確認する必要がある",
            "直接条件が見つかった場合は試薬量、mmol、溶媒、温度、時間、精製、収率、NMR、DOI、SIリンクを残す",
        ],
        "output": """# 2-Butoxynaphthalene 合成条件

出典:
- Nature Communications 2024
- Electrochemical on-surface synthesis of a strong electron-donating graphene nanoribbon catalyst
- DOI: 10.1038/s41467-024-50086-6
- Supplementary Information: https://static-content.springer.com/esm/art%3A10.1038%2Fs41467-024-50086-6/MediaObjects/41467_2024_50086_MOESM1_ESM.pdf

条件:
- 2-naphthol: 2.88 g, 20.0 mmol
- 1-bromobutane: 3.30 g, 24.1 mmol
- KOH: 2.25 g, 40.1 mmol
- DMSO: 30 mL
- 95 degC, overnight

後処理/精製:
- 水で希釈
- 冷蔵庫で冷却
- 析出固体をろ過
- silica gel chromatography, hexanes
- GPC, CHCl3

収量:
- 2.88 g, 14.4 mmol, 72%
- white solid

教訓:
- 有機合成Agentは本文だけでなくSI/ESI/PDF/特許まで検索する。
""",
        "metadata": {
            "doi": "10.1038/s41467-024-50086-6",
            "si_url": "https://static-content.springer.com/esm/art%3A10.1038%2Fs41467-024-50086-6/MediaObjects/41467_2024_50086_MOESM1_ESM.pdf",
        },
    },
    {
        "agent": "検証Agent",
        "kind": "summaries",
        "request": "検証Agentの仕事内容定義",
        "decision_summary": "各Agentの成果物を信じてよい状態か判定する係として定義した。",
        "verification": [
            "終了コードを確認する",
            "Traceback/Exception/失敗語を確認する",
            "OneNoteリンク、DOI、SI確認、分類件数、Markdown保存などAgent別の成果物を確認する",
            "結果をAgent_検証AgentとMarkdownへ残す",
        ],
        "output": "検証Agentは、実行後に成功/警告/失敗を判定し、警告または要対応事項を明示する。",
        "metadata": {},
    },
]


def backfill_route_log() -> list[Path]:
    written: list[Path] = []
    if not ROUTE_LOG.exists():
        return written

    base_time = datetime.fromtimestamp(ROUTE_LOG.stat().st_mtime) - timedelta(minutes=30)
    with ROUTE_LOG.open(encoding="utf-8") as f:
        for index, line in enumerate(f):
            line = line.strip()
            if not line:
                continue
            record = json.loads(line)
            created_at = base_time + timedelta(seconds=index)
            agent = record.get("agent", "UnknownAgent")
            request = record.get("request", "unknown request")
            written.append(
                write_markdown_record(
                    agent="司令塔Agent",
                    kind="logs",
                    request=request,
                    decision_summary=f"{agent} に割り振りました。理由: {record.get('reason', '')}",
                    command=record.get("command", []),
                    verification=[
                        "依頼内容を解釈する",
                        "担当Agentを選定する",
                        "担当Agentの実行コマンドと検証項目を記録する",
                    ],
                    output=f"exit_code={record.get('exit_code')}",
                    metadata=record,
                    created_at=created_at,
                )
            )
            written.append(
                write_markdown_record(
                    agent=agent,
                    kind="logs",
                    request=request,
                    decision_summary=record.get("reason", ""),
                    command=record.get("command", []),
                    verification=record.get("verification", []),
                    output=f"exit_code={record.get('exit_code')}",
                    metadata=record,
                    created_at=created_at + timedelta(milliseconds=500),
                )
            )
            written.append(
                write_json_record(
                    agent=agent,
                    kind="route_record",
                    request=request,
                    data=record,
                    created_at=created_at + timedelta(milliseconds=700),
                )
            )
    return written


def backfill_manual() -> list[Path]:
    written: list[Path] = []
    now = datetime.now() - timedelta(minutes=5)
    for index, record in enumerate(MANUAL_BACKFILL):
        created_at = now + timedelta(seconds=index)
        written.append(
            write_markdown_record(
                agent=record["agent"],
                kind=record["kind"],
                request=record["request"],
                decision_summary=record["decision_summary"],
                command=[],
                verification=record["verification"],
                output=record["output"],
                metadata=record["metadata"],
                created_at=created_at,
            )
        )
    return written


def main() -> int:
    written = []
    written.extend(backfill_route_log())
    written.extend(backfill_manual())
    print(f"created_files={len(written)}")
    for path in written:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
