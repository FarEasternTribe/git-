from __future__ import annotations

import argparse
import json
import os
import platform
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from agent_file_logger import write_markdown_record


WORKSPACE_DIR = Path(__file__).resolve().parent
DEFAULT_NOTEBOOK = os.getenv("AGENT_LOG_NOTEBOOK", "OpenAI_agent1")
DEFAULT_PDF_DIRS = [WORKSPACE_DIR / "paper", WORKSPACE_DIR / "papers"]
COMMAND_LOG_SECTION = "命令したLog"


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


def device_prefix() -> str:
    return f"[{detect_device_label()}]"


def ps_single_quoted(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def escape_cdata(text: str) -> str:
    return text.replace("]]>", "]]]]><![CDATA[>")


def find_recent_pdfs(limit: int = 10) -> list[Path]:
    pdfs: list[Path] = []
    for root in DEFAULT_PDF_DIRS:
        if root.exists():
            pdfs.extend(root.rglob("*.pdf"))
    pdfs.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return pdfs[:limit]


def build_page_lines(
    *,
    agent: str,
    request: str,
    decision_summary: str,
    command: list[str],
    verification: list[str],
    output: str,
    pdfs: list[Path],
) -> list[str]:
    now = datetime.now()
    lines = [
        f"Agent: {agent}",
        f"日付: {now.strftime('%Y-%m-%d')}",
        f"時刻: {now.strftime('%H:%M:%S')}",
        f"記録日時: {now.isoformat(timespec='seconds')}",
        f"依頼: {request}",
        "",
        "判断概要:",
        decision_summary or "記録なし",
        "",
        "実行コマンド:",
        subprocess.list2cmdline(command) if command else "なし",
        "",
        "検証項目:",
    ]
    lines.extend(f"- {item}" for item in verification)
    lines.extend(["", "実行出力/結果概要:", output.strip() or "なし"])
    if pdfs:
        lines.extend(["", "関連PDF:"])
        lines.extend(f"- {pdf.name}: {pdf}" for pdf in pdfs)
    return lines


def build_powershell_script(
    *,
    notebook: str,
    section: str,
    title: str,
    lines: list[str],
    pdfs: list[Path],
) -> str:
    line_xml = "\n".join(
        f"""      <one:OE><one:T><![CDATA[{escape_cdata(line)}]]></one:T></one:OE>"""
        for line in lines
    )
    pdf_xml = "\n".join(
        f"""
    <one:InsertedFile pathCache={ps_single_quoted(str(pdf))} preferredName={ps_single_quoted(pdf.name)}>
      <one:Position x="36" y="{160 + index * 40}" z="0"/>
    </one:InsertedFile>"""
        for index, pdf in enumerate(pdfs)
    )

    page_xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="__PAGE_ID__">
  <one:Title>
    <one:OE>
      <one:T><![CDATA[{escape_cdata(title)}]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0"/>
    <one:OEChildren>
{line_xml}
    </one:OEChildren>
  </one:Outline>
{pdf_xml}
</one:Page>"""

    return f"""
$ErrorActionPreference = 'Stop'
$NotebookName = {ps_single_quoted(notebook)}
$SectionName = {ps_single_quoted(section)}
$PageXmlTemplate = @'
{page_xml}
'@

$one = New-Object -ComObject OneNote.Application
[xml]$hierarchy = ''
$one.GetHierarchy('', 4, [ref]$hierarchy)

$notebook = $hierarchy.DocumentElement.SelectNodes('//*') |
  Where-Object {{ $_.LocalName -eq 'Notebook' -and $_.name -eq $NotebookName }} |
  Select-Object -First 1
if ($null -eq $notebook) {{ throw "Notebook not found: $NotebookName" }}

$section = $notebook.SelectNodes('.//*[local-name()="Section"]') |
  Where-Object {{ $_.name -eq $SectionName }} |
  Select-Object -First 1

if ($null -eq $section) {{
  $ns = $notebook.NamespaceURI
  $section = $hierarchy.CreateElement('one', 'Section', $ns)
  $section.SetAttribute('name', $SectionName)
  $section.SetAttribute('path', $notebook.path + $SectionName + '.one')
  $firstSectionGroup = $notebook.SelectNodes('./*[local-name()="SectionGroup"]') | Select-Object -First 1
  if ($null -ne $firstSectionGroup) {{
    [void]$notebook.InsertBefore($section, $firstSectionGroup)
  }} else {{
    [void]$notebook.AppendChild($section)
  }}
  $one.UpdateHierarchy($notebook.OuterXml)
  Start-Sleep -Milliseconds 800
  [xml]$hierarchy = ''
  $one.GetHierarchy('', 4, [ref]$hierarchy)
  $notebook = $hierarchy.DocumentElement.SelectNodes('//*') |
    Where-Object {{ $_.LocalName -eq 'Notebook' -and $_.name -eq $NotebookName }} |
    Select-Object -First 1
  $section = $notebook.SelectNodes('.//*[local-name()="Section"]') |
    Where-Object {{ $_.name -eq $SectionName }} |
    Select-Object -First 1
}}
if ($null -eq $section) {{ throw "Section not found or could not be created: $SectionName" }}

$pageId = ''
$one.CreateNewPage($section.ID, [ref]$pageId, 0)
$pageXml = $PageXmlTemplate.Replace('__PAGE_ID__', $pageId)
try {{
  $one.UpdatePageContent($pageXml)
  "created $pageId"
}} catch {{
  $fallbackXml = $pageXml -replace '<one:InsertedFile[\\s\\S]*?</one:InsertedFile>', ''
  $one.UpdatePageContent($fallbackXml)
  "created_without_pdf_attachments $pageId"
  "attachment_error $($_.Exception.Message)"
}}
"""


def write_agent_log(
    *,
    agent: str,
    request: str,
    decision_summary: str,
    command: list[str],
    verification: list[str],
    output: str,
    notebook: str = DEFAULT_NOTEBOOK,
    section_name: str | None = None,
    attach_pdfs: bool = False,
) -> tuple[bool, str]:
    markdown_path = write_markdown_record(
        agent=agent,
        kind="logs",
        request=request,
        decision_summary=decision_summary,
        command=command,
        verification=verification,
        output=output,
        metadata={
            "notebook": notebook,
            "attach_pdfs": attach_pdfs,
        },
    )
    section = section_name or f"Agent_{agent}"
    command_log_device_prefix = device_prefix()
    title_prefix = f"{command_log_device_prefix} " if section == COMMAND_LOG_SECTION else ""
    title = f"{title_prefix}{datetime.now().strftime('%Y-%m-%d_%H%M%S')} {agent}"
    pdfs = find_recent_pdfs() if attach_pdfs else []
    lines = build_page_lines(
        agent=agent,
        request=request,
        decision_summary=decision_summary,
        command=command,
        verification=verification,
        output=output,
        pdfs=pdfs,
    )
    if section == COMMAND_LOG_SECTION:
        lines.insert(0, command_log_device_prefix)
    script = build_powershell_script(
        notebook=notebook,
        section=section,
        title=title,
        lines=lines,
        pdfs=pdfs,
    )
    with tempfile.TemporaryDirectory(prefix="agent_onenote_log_") as tmp:
        ps_path = Path(tmp) / "write_agent_log.ps1"
        ps_path.write_text(script, encoding="utf-8-sig")
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_path)],
            cwd=WORKSPACE_DIR,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    detail = "\n".join(part for part in [result.stdout.strip(), result.stderr.strip()] if part)
    detail = f"{detail}\nmarkdown_log={markdown_path}".strip()
    return result.returncode == 0, detail


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Agentの判断概要・実行内容・検証結果をOneNoteへ記録します。")
    parser.add_argument("--agent", required=True)
    parser.add_argument("--request", required=True)
    parser.add_argument("--decision-summary", default="")
    parser.add_argument("--command-json", default="[]")
    parser.add_argument("--verification-json", default="[]")
    parser.add_argument("--output", default="")
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--section", default="")
    parser.add_argument("--attach-pdfs", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    command = json.loads(args.command_json)
    verification = json.loads(args.verification_json)
    ok, detail = write_agent_log(
        agent=args.agent,
        request=args.request,
        decision_summary=args.decision_summary,
        command=command,
        verification=verification,
        output=args.output,
        notebook=args.notebook,
        section_name=args.section or None,
        attach_pdfs=args.attach_pdfs,
    )
    print(detail)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
