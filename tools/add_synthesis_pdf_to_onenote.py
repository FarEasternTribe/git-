from __future__ import annotations

import argparse
import html
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parents[1]
if str(WORKSPACE_DIR) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_DIR))

from agent_common import escape_cdata, ps_single_quoted

DEFAULT_NOTEBOOK = "2026実験"
DEFAULT_SECTION = "有機合成"


@dataclass
class ExtractionResult:
    doi_list: list[str]
    title_guess: str
    experimental_snippets: list[str]
    condition_rows: list[tuple[str, str]]
    text_char_count: int


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_pdf_text(pdf_path: Path, max_pages: int | None = None) -> str:
    try:
        from pypdf import PdfReader
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("pypdf is required for PDF text extraction") from exc

    reader = PdfReader(str(pdf_path))
    pages = reader.pages[:max_pages] if max_pages else reader.pages
    parts: list[str] = []
    for index, page in enumerate(pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            text = f"[page {index}: text extraction failed: {exc}]"
        if text.strip():
            parts.append(f"\n--- page {index} ---\n{text}")
    return normalize_text("\n".join(parts))


def extract_dois(text: str) -> list[str]:
    pattern = r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b"
    values = []
    for match in re.finditer(pattern, text, flags=re.IGNORECASE):
        doi = match.group(0).rstrip(".,;)")
        if doi not in values:
            values.append(doi)
    return values


def guess_title(text: str, fallback: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    candidates = []
    for line in lines[:80]:
        if 12 <= len(line) <= 180 and not re.search(r"^(abstract|introduction|references?)$", line, re.I):
            if not re.match(r"^(doi|https?://|www\.|received|published|copyright)", line, re.I):
                candidates.append(line)
    return candidates[0] if candidates else fallback


def window_around(text: str, start: int, end: int, radius: int = 650) -> str:
    left = max(0, start - radius)
    right = min(len(text), end + radius)
    snippet = text[left:right]
    return re.sub(r"\s+", " ", snippet).strip()


def extract_experimental_snippets(text: str, limit: int = 8) -> list[str]:
    keywords = [
        r"experimental",
        r"general procedure",
        r"synthesis of",
        r"preparation of",
        r"was prepared",
        r"was synthesized",
        r"compound\s+\d+",
        r"yield",
        r"mmol",
        r"equiv",
        r"under nitrogen",
        r"under argon",
        r"chromatograph",
        r"recrystall",
    ]
    snippets: list[str] = []
    for pattern in keywords:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            snippet = window_around(text, match.start(), match.end())
            if snippet and all(snippet[:160] not in existing for existing in snippets):
                snippets.append(snippet)
            if len(snippets) >= limit:
                return snippets
    return snippets


def find_first(patterns: list[str], text: str) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return re.sub(r"\s+", " ", match.group(0)).strip()
    return "未抽出"


def build_condition_rows(snippets: list[str], doi_list: list[str]) -> list[tuple[str, str]]:
    joined = " ".join(snippets)
    rows = [
        ("根拠DOI", ", ".join(doi_list) if doi_list else "未抽出"),
        (
            "基質/出発物質",
            find_first(
                [
                    r"C3N3F3|cyanuric fluoride|cyanuric chloride|2,4,6-trichloro-1,3,5-triazine|2,4,6-trifluoro-1,3,5-triazine",
                    r"[A-Z][A-Za-z0-9,\-\(\)\[\] ]{3,80}(?:\s*\(\s*\d+(?:\.\d+)?\s*(?:g|mg|mmol)\s*\))",
                ],
                joined,
            ),
        ),
        (
            "アルキン/求核剤",
            find_first(
                [
                    r"(?:\(trimethylsilyl\)ethynyl|trimethylsilylethynyl|ethynyl)zinc chloride",
                    r"(?:\(trimethylsilyl\)ethynyl|trimethylsilylethynyl|ethynyl)zinc bromide",
                    r"(?:\(trimethylsilyl\)ethynyl|trimethylsilylethynyl|ethynyl)zinc [a-z]+",
                    r"(?:lithium|sodium|potassium)[-\s]+(?:\(trimethylsilyl\))?acetylide",
                    r"alkali[-\s]+metal\s+\(trimethylsilyl\)acetylides?",
                    r"trimethylsilylacetylene|TMS[-\s]*acetylene|\(trimethylsilyl\)acetylene",
                ],
                joined,
            ),
        ),
        (
            "触媒",
            find_first(
                [
                    r"Pd\s*\(\s*PPh\s*3\s*\)\s*4",
                    r"Pd\s*\(\s*PPh\s*3\s*\)\s*2\s*Cl\s*2",
                    r"PdCl2",
                    r"palladium catalyst",
                    r"CuI",
                    r"copper\(I\) iodide",
                ],
                joined,
            ),
        ),
        (
            "塩基",
            find_first([r"triethylamine|Et3N|diisopropylamine|i-Pr2NH|K2CO3|n-BuLi|BuLi"], joined),
        ),
        (
            "溶媒",
            find_first([r"THF|tetrahydrofuran|ether|diethyl ether|toluene|benzene|DMF|DMSO|CH2Cl2|dichloromethane"], joined),
        ),
        (
            "温度",
            find_first([r"-?\d+\s*°?\s*C|room temperature|rt|reflux"], joined),
        ),
        (
            "時間",
            find_first([r"\d+(?:\.\d+)?\s*(?:h|hr|hours?|min|minutes?)|overnight"], joined),
        ),
        (
            "収率",
            find_first([r"\d+(?:\.\d+)?\s*%"], joined),
        ),
        (
            "精製",
            find_first([r"column chromatography|chromatograph[a-z]*|recrystall[a-z]*|sublim[a-z]*|distill[a-z]*|washed|filtered"], joined),
        ),
    ]
    return rows


def extract_synthesis_info(pdf_path: Path, max_pages: int | None = None) -> ExtractionResult:
    text = extract_pdf_text(pdf_path, max_pages=max_pages)
    doi_list = extract_dois(text)
    snippets = extract_experimental_snippets(text)
    return ExtractionResult(
        doi_list=doi_list,
        title_guess=guess_title(text, pdf_path.stem),
        experimental_snippets=snippets,
        condition_rows=build_condition_rows(snippets, doi_list),
        text_char_count=len(text),
    )


def build_lines(pdf_path: Path, page_title: str, extraction: ExtractionResult) -> list[str]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"追加日時: {now}",
        f"PDF: {pdf_path.name}",
        f"抽出タイトル候補: {extraction.title_guess}",
        f"抽出文字数: {extraction.text_char_count}",
        "",
        "合成条件表（PDF自動抽出・要原文確認）",
    ]
    for key, value in extraction.condition_rows:
        lines.append(f"| {key} | {value} |")

    lines.extend(["", "DOI候補:"])
    if extraction.doi_list:
        lines.extend(f"- {doi}" for doi in extraction.doi_list)
    else:
        lines.append("- 未抽出")

    lines.extend(["", "実験項候補抜粋（自動抽出）:"])
    if extraction.experimental_snippets:
        for index, snippet in enumerate(extraction.experimental_snippets, start=1):
            lines.append(f"[候補{index}] {snippet[:1800]}")
    else:
        lines.append("実験項らしいテキストを抽出できませんでした。PDFのOCRまたは本文確認が必要です。")

    lines.extend(
        [
            "",
            "未確認点:",
            "- 自動抽出のため、当量・収率・化合物番号・スペクトル値は原文PDFで確認してください。",
            "- DOIが未抽出の場合はCrossref/出版社ページ/SI/PDF本文を再検索してください。",
        ]
    )
    return lines


def build_powershell_script(
    *,
    pdf_path: Path,
    notebook: str,
    section: str,
    page_title: str,
    lines: list[str],
    append_existing: bool,
) -> str:
    line_xml = "\n".join(
        f"""      <one:OE><one:T><![CDATA[{escape_cdata(line)}]]></one:T></one:OE>"""
        for line in lines
    )
    escaped_pdf = html.escape(str(pdf_path), quote=True)
    escaped_name = html.escape(pdf_path.name, quote=True)
    attach_xml = f"""      <one:OE><one:InsertedFile pathCache="{escaped_pdf}" preferredName="{escaped_name}"/></one:OE>"""
    page_xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="__PAGE_ID__">
  <one:Title><one:OE><one:T><![CDATA[{escape_cdata(page_title)}]]></one:T></one:OE></one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0"/>
    <one:OEChildren>
{line_xml}
{attach_xml}
    </one:OEChildren>
  </one:Outline>
</one:Page>"""

    append_xml = f"""<one:Outline xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
    <one:Position x="36" y="220" z="1"/>
    <one:OEChildren>
{line_xml}
    </one:OEChildren>
  </one:Outline>"""

    return f"""
$ErrorActionPreference = 'Stop'
$NotebookName = {ps_single_quoted(notebook)}
$SectionName = {ps_single_quoted(section)}
$PageTitle = {ps_single_quoted(page_title)}
$AppendExisting = ${str(append_existing).lower()}
$PageXmlTemplate = @'
{page_xml}
'@
$AppendXml = @'
{append_xml}
'@

$one = New-Object -ComObject OneNote.Application
[string]$hierXml = ''
$one.GetHierarchy('', 4, [ref]$hierXml)
[xml]$hier = $hierXml
$ns = New-Object System.Xml.XmlNamespaceManager($hier.NameTable)
$ns.AddNamespace('one','http://schemas.microsoft.com/office/onenote/2013/onenote')
$notebook = @($hier.SelectNodes('//one:Notebook', $ns)) | Where-Object {{ $_.name -eq $NotebookName }} | Select-Object -First 1
if (-not $notebook) {{ throw "Notebook not found: $NotebookName" }}
$sectionNode = @($notebook.SelectNodes('.//one:Section', $ns)) | Where-Object {{ $_.name -eq $SectionName }} | Select-Object -First 1
if (-not $sectionNode) {{ throw "Section not found: $SectionName" }}

[string]$sectionXml = ''
$one.GetHierarchy($sectionNode.ID, 4, [ref]$sectionXml)
[xml]$secDoc = $sectionXml
$existing = @($secDoc.SelectNodes('//*[local-name()="Page"]')) | Where-Object {{ $_.name -eq $PageTitle }} | Select-Object -First 1

if ($existing -and $AppendExisting) {{
  $pageId = $existing.ID
  [string]$pageXml = ''
  $one.GetPageContent($pageId, [ref]$pageXml)
  [xml]$pageDoc = $pageXml
  $fragment = $pageDoc.CreateDocumentFragment()
  $fragment.InnerXml = $AppendXml
  [void]$pageDoc.DocumentElement.AppendChild($fragment)
  $one.UpdatePageContent($pageDoc.OuterXml)
  $mode = 'updated_existing'
}} else {{
  [string]$pageId = ''
  $one.CreateNewPage($sectionNode.ID, [ref]$pageId, 0)
  $one.UpdatePageContent($PageXmlTemplate.Replace('__PAGE_ID__', $pageId))
  $mode = 'created'
}}

Start-Sleep -Seconds 1
[string]$readXml = ''
$one.GetPageContent($pageId, [ref]$readXml)
$link = ''
try {{ $one.GetHyperlinkToObject($pageId, '', [ref]$link) }} catch {{ $link = '' }}
"mode=$mode"
"page_id=$pageId"
"onenote_link=$link"
"title_ok=$($readXml.Contains($PageTitle))"
"doi_section_ok=$($readXml.Contains('DOI候補'))"
"table_ok=$($readXml.Contains('合成条件表'))"
"file_name_ok=$($readXml.Contains({ps_single_quoted(pdf_path.name)}))"
"""


def write_to_onenote(
    *,
    pdf_path: Path,
    notebook: str,
    section: str,
    page_title: str,
    lines: list[str],
    append_existing: bool,
) -> tuple[int, str]:
    script = build_powershell_script(
        pdf_path=pdf_path,
        notebook=notebook,
        section=section,
        page_title=page_title,
        lines=lines,
        append_existing=append_existing,
    )
    with tempfile.TemporaryDirectory(prefix="synthesis_pdf_onenote_") as tmp:
        ps_path = Path(tmp) / "write_synthesis_pdf.ps1"
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
    output = "\n".join(part for part in [result.stdout.strip(), result.stderr.strip()] if part)
    return result.returncode, output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="有機合成PDFをOneNoteへ添付し、PDF本文から実験項候補と合成条件表を同じページに追記します。"
    )
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--section", default=DEFAULT_SECTION)
    parser.add_argument("--title", help="OneNoteページタイトル。省略時はPDFファイル名から生成")
    parser.add_argument("--append-existing", action="store_true", help="同名ページがあれば追記する")
    parser.add_argument("--max-pages", type=int, help="抽出対象ページ数。省略時は全ページ")
    parser.add_argument("--extract-only", action="store_true", help="OneNoteへ書き込まず抽出結果だけ表示")
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    args = parse_args()
    pdf_path = args.pdf.expanduser().resolve()
    if not pdf_path.exists():
        print(f"PDF not found: {pdf_path}", file=sys.stderr)
        return 2

    page_title = args.title or pdf_path.stem
    extraction = extract_synthesis_info(pdf_path, max_pages=args.max_pages)
    lines = build_lines(pdf_path, page_title, extraction)

    print(f"PDF: {pdf_path}")
    print(f"Title: {page_title}")
    print(f"TextChars: {extraction.text_char_count}")
    print("DOI: " + (", ".join(extraction.doi_list) if extraction.doi_list else "未抽出"))
    print("ConditionTable:")
    for key, value in extraction.condition_rows:
        print(f"- {key}: {value}")

    if args.extract_only:
        return 0

    code, output = write_to_onenote(
        pdf_path=pdf_path,
        notebook=args.notebook,
        section=args.section,
        page_title=page_title,
        lines=lines,
        append_existing=args.append_existing,
    )
    print(output)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
