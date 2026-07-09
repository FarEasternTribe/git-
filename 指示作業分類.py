from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path


DEFAULT_NOTEBOOK = "2025年書込テスト"
DEFAULT_MAP_FILE = Path("tools") / "onenote_classification_map.json"

CATEGORY_RULES = [
    {
        "key": "inbox",
        "section": "分類_01_未整理・一時置き場",
        "patterns": [
            "無題",
            "手書き",
            "メモ",
            "GPT",
            "AI",
            "自己分析",
            "ライフハック",
            "勉強法",
        ],
    },
    {
        "key": "research",
        "section": "分類_02_研究・論文・技術メモ",
        "patterns": [
            "GNR",
            "グラフェン",
            "ジメトキシ",
            "Br化",
            "QD",
            "量子ドット",
            "GeSe",
            "強誘電",
            "シフト電流",
            "合成",
            "特許",
            "patent",
            "espacenet",
            "TEM",
            "600 mesh",
            "ACS",
            "論文",
            "化学",
            "半導体",
            "研究概要",
        ],
    },
    {
        "key": "work",
        "section": "分類_03_仕事・大学業務",
        "patterns": [
            "R7年度",
            "センター談話会",
            "作業環境測定",
            "松田先生",
            "中嶋先生",
            "京都大学",
            "坂口研究室",
            "シラバス",
            "講習",
            "依頼",
            "文科省",
            "申請",
        ],
    },
    {
        "key": "writing",
        "section": "分類_04_文章・メール・発表テンプレ",
        "patterns": [
            "メール",
            "テンプレ",
            "講演者各位",
            "Members,",
            "Hello, my name",
            "プレゼントーク",
            "Slide",
            "お世話になっております",
            "自己紹介",
        ],
    },
    {
        "key": "journal",
        "section": "分類_05_日誌・思考整理・ToDo",
        "patterns": [
            "Timeline",
            "タイムライン",
            "ToDo",
            "To-Do",
            "対話まとめ",
            "日誌",
            "ジャーナリング",
            "問題発生時チェックシート",
            "思考",
            "要約サマリー",
            "最新ルール",
        ],
    },
    {
        "key": "life",
        "section": "分類_06_生活・個人メモ",
        "patterns": [
            "料理",
            "レシピ",
            "delishkitchen",
            "オランデーズ",
            "麻醤",
            "楽天",
            "トナー",
            "車",
            "慢性疲労",
            "食べたい",
            "買い回り",
        ],
    },
    {
        "key": "web",
        "section": "分類_07_Webクリップ・あとで読む",
        "patterns": [
            "http://",
            "https://",
            "note.com",
            "SmartNews",
            "kakaku",
            "Amazon.co.jp",
            "Lifehacker",
            "ライフハッカー",
            "YouTube",
            "コンテンツへとスキップ",
        ],
    },
]


def ps_single_quoted(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def build_powershell_script(
    *,
    notebook: str,
    execute: bool,
    reset_classification_pages: bool,
    map_file: Path,
) -> str:
    categories = []
    for category in CATEGORY_RULES:
        patterns = ", ".join(ps_single_quoted(pattern) for pattern in category["patterns"])
        categories.append(
            "@{ Key = "
            + ps_single_quoted(category["key"])
            + "; Section = "
            + ps_single_quoted(category["section"])
            + "; Patterns = @("
            + patterns
            + ") }"
        )
    categories_ps = ",\n  ".join(categories)
    execute_ps = "$true" if execute else "$false"
    reset_ps = "$true" if reset_classification_pages else "$false"
    map_file_ps = ps_single_quoted(str(map_file))

    return f"""
$ErrorActionPreference = 'Stop'
$NotebookName = {ps_single_quoted(notebook)}
$Execute = {execute_ps}
$ResetClassificationPages = {reset_ps}
$MapPath = {map_file_ps}

$Categories = @(
  {categories_ps}
)

function Normalize-Text([string]$Text) {{
  if ($null -eq $Text) {{ return '' }}
  return ($Text -replace '\\s+', ' ').Trim().ToLowerInvariant()
}}

function Escape-CData([string]$Text) {{
  if ($null -eq $Text) {{ return '' }}
  return $Text.Replace(']]>', ']]]]><![CDATA[>')
}}

function Refresh-Hierarchy($OneNote) {{
  [xml]$hierarchy = ''
  $OneNote.GetHierarchy('', 4, [ref]$hierarchy)
  return $hierarchy
}}

function Get-Notebook($Hierarchy, [string]$Name) {{
  return $Hierarchy.DocumentElement.SelectNodes('//*') |
    Where-Object {{ $_.LocalName -eq 'Notebook' -and $_.name -eq $Name }} |
    Select-Object -First 1
}}

function Select-Category($Page, [string]$SourceSectionName) {{
  $target = (Normalize-Text ($SourceSectionName + ' ' + $Page.name))
  foreach ($category in $Categories) {{
    foreach ($pattern in $category.Patterns) {{
      if ($target.Contains((Normalize-Text $pattern))) {{
        return $category
      }}
    }}
  }}
  return ($Categories | Where-Object {{ $_.Key -eq 'inbox' }} | Select-Object -First 1)
}}

function Ensure-Section($OneNote, [xml]$Hierarchy, $Notebook, [string]$SectionName) {{
  $existing = $Notebook.SelectNodes('.//*[local-name()="Section"]') |
    Where-Object {{ $_.name -eq $SectionName }} |
    Select-Object -First 1
  if ($null -ne $existing) {{ return $existing.ID }}

  $ns = $Notebook.NamespaceURI
  $section = $Hierarchy.CreateElement('one', 'Section', $ns)
  $section.SetAttribute('name', $SectionName)
  $section.SetAttribute('path', $Notebook.path + $SectionName + '.one')

  $firstSectionGroup = $Notebook.SelectNodes('./*[local-name()="SectionGroup"]') | Select-Object -First 1
  if ($null -ne $firstSectionGroup) {{
    [void]$Notebook.InsertBefore($section, $firstSectionGroup)
  }} else {{
    [void]$Notebook.AppendChild($section)
  }}

  $OneNote.UpdateHierarchy($Notebook.OuterXml)
  Start-Sleep -Milliseconds 800
  $updated = Refresh-Hierarchy $OneNote
  $updatedNotebook = Get-Notebook $updated $NotebookName
  $created = $updatedNotebook.SelectNodes('.//*[local-name()="Section"]') |
    Where-Object {{ $_.name -eq $SectionName }} |
    Select-Object -First 1
  if ($null -eq $created) {{ throw "Failed to create section: $SectionName" }}
  return $created.ID
}}

function New-PageXml([string]$Title, [string[]]$Lines) {{
  $oeParts = New-Object System.Collections.ArrayList
  foreach ($line in $Lines) {{
    if ([string]::IsNullOrWhiteSpace($line)) {{ continue }}
    [void]$oeParts.Add(@"
      <one:OE>
        <one:T><![CDATA[$(Escape-CData $line)]]></one:T>
      </one:OE>
"@)
  }}
  $body = ($oeParts -join "`n")
  return @"
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="__PAGE_ID__">
  <one:Title>
    <one:OE>
      <one:T><![CDATA[$(Escape-CData $Title)]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0"/>
    <one:OEChildren>
$body
    </one:OEChildren>
  </one:Outline>
</one:Page>
"@
}}

function New-ClassificationCard($OneNote, [string]$TargetSectionId, $SourcePage, [string]$SourceSectionName, [string]$CategorySectionName) {{
  $pageId = ''
  $OneNote.CreateNewPage($TargetSectionId, [ref]$pageId, 0)

  $link = ''
  try {{
    $OneNote.GetHyperlinkToObject($SourcePage.ID, '', [ref]$link)
  }} catch {{
    $link = $SourcePage.ID
  }}

  $title = $SourcePage.name
  if ([string]::IsNullOrWhiteSpace($title)) {{ $title = '無題のページ' }}
  $lines = @(
    "分類: $CategorySectionName",
    "元セクション: $SourceSectionName",
    "元ページ: $title",
    "元ページリンク: $link",
    "注記: 元ノートは削除・移動していません。このページは分類用の参照カードです。"
  )
  $pageXml = (New-PageXml $title $lines).Replace('__PAGE_ID__', $pageId)
  $OneNote.UpdatePageContent($pageXml)
  return $pageId
}}

$one = New-Object -ComObject OneNote.Application
$hierarchy = Refresh-Hierarchy $one
$notebook = Get-Notebook $hierarchy $NotebookName
if ($null -eq $notebook) {{ throw "Notebook not found: $NotebookName" }}

$sourceSections = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {{
  $_.name -notlike '分類_*' -and $_.name -ne '削除されたページ'
}})

$plan = New-Object System.Collections.ArrayList
foreach ($section in $sourceSections) {{
  foreach ($page in @($section.SelectNodes('.//*[local-name()="Page"]'))) {{
    $category = Select-Category $page $section.name
    [void]$plan.Add([pscustomobject]@{{
      SourceSection = $section.name
      SourcePageId = $page.ID
      Title = $page.name
      CategoryKey = $category.Key
      CategorySection = $category.Section
    }})
  }}
}}

$summary = @{{}}
foreach ($category in $Categories) {{ $summary[$category.Section] = 0 }}
foreach ($item in $plan) {{ $summary[$item.CategorySection] += 1 }}

"Notebook: $NotebookName"
"対象元セクション: $($sourceSections.Count)"
"対象ページ: $($plan.Count)"
foreach ($category in $Categories) {{
  "$($category.Section): $($summary[$category.Section])"
}}

if (-not $Execute) {{
  "DRY-RUN: OneNoteには書き込んでいません。実行するには --execute を付けてください。"
  exit 0
}}

$categorySectionIds = @{{}}
foreach ($category in $Categories) {{
  $categorySectionIds[$category.Key] = Ensure-Section $one $hierarchy $notebook $category.Section
}}

$hierarchy = Refresh-Hierarchy $one
$notebook = Get-Notebook $hierarchy $NotebookName
$targetSections = @{{}}
foreach ($category in $Categories) {{
  $targetSections[$category.Key] = $notebook.SelectNodes('.//*[local-name()="Section"]') |
    Where-Object {{ $_.name -eq $category.Section }} |
    Select-Object -First 1
}}

if ($ResetClassificationPages) {{
  foreach ($category in $Categories) {{
    foreach ($page in @($targetSections[$category.Key].SelectNodes('.//*[local-name()="Page"]'))) {{
      $one.DeleteHierarchy($page.ID)
    }}
  }}
}}

$created = @{{}}
foreach ($category in $Categories) {{ $created[$category.Key] = 0 }}
$failed = New-Object System.Collections.ArrayList
$entries = New-Object System.Collections.ArrayList

foreach ($item in $plan) {{
  try {{
    $targetId = New-ClassificationCard $one $categorySectionIds[$item.CategoryKey] $item $item.SourceSection $item.CategorySection
    [void]$entries.Add([pscustomobject]@{{
      sourcePageId = $item.SourcePageId
      sourceSection = $item.SourceSection
      title = $item.Title
      category = $item.CategoryKey
      targetPageId = $targetId
      classifiedAt = (Get-Date).ToString('o')
    }})
    $created[$item.CategoryKey] += 1
  }} catch {{
    [void]$failed.Add([pscustomobject]@{{
      sourceSection = $item.SourceSection
      title = $item.Title
      category = $item.CategoryKey
      error = $_.Exception.Message
    }})
  }}
}}

$map = [pscustomobject]@{{
  notebook = $NotebookName
  updatedAt = (Get-Date).ToString('o')
  entries = @($entries)
}}
$mapDir = Split-Path -Parent $MapPath
if ($mapDir -and -not (Test-Path -LiteralPath $mapDir)) {{
  New-Item -ItemType Directory -Path $mapDir | Out-Null
}}
$map | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $MapPath -Encoding UTF8

"実行結果:"
foreach ($category in $Categories) {{
  "$($category.Section): 作成 $($created[$category.Key])"
}}
"失敗: $($failed.Count)"
foreach ($item in $failed) {{
  "FAILED: [$($item.sourceSection)] $($item.title) -> $($item.category): $($item.error)"
}}
"""


def run_powershell(script: str) -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory(prefix="onenote_classify_") as tmp:
        ps_path = Path(tmp) / "classify_onenote.ps1"
        ps_path.write_text(script, encoding="utf-8-sig")
        return subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ps_path),
            ],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "OneNoteノートブック内の各セクションのページを分類し、"
            "分類セクションに元ページリンク付きカードを作成します。"
        )
    )
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK, help="分類対象のOneNoteノートブック名")
    parser.add_argument("--execute", action="store_true", help="OneNoteへ分類セクション/分類カードを書き込む")
    parser.add_argument(
        "--keep-existing",
        action="store_true",
        help="既存の分類セクション内ページを消さずに追記する。未指定時は分類セクション内だけ作り直す",
    )
    parser.add_argument("--map-file", type=Path, default=DEFAULT_MAP_FILE, help="分類結果JSONの保存先")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    script = build_powershell_script(
        notebook=args.notebook,
        execute=args.execute,
        reset_classification_pages=not args.keep_existing,
        map_file=args.map_file,
    )
    result = run_powershell(script)
    if result.stdout:
        print(result.stdout.rstrip())
    if result.stderr:
        print(result.stderr.rstrip(), file=sys.stderr)
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
