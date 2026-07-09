from __future__ import annotations

import argparse
import difflib
import re
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path


DEFAULT_MAP_DIR = Path("tools") / "onenote_classification_maps"

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
    incremental: bool,
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
    incremental_ps = "$true" if incremental else "$false"

    return f"""
$ErrorActionPreference = 'Stop'
$NotebookName = {ps_single_quoted(notebook)}
$Execute = {execute_ps}
$ResetClassificationPages = {reset_ps}
$MapPath = {map_file_ps}
$Incremental = {incremental_ps}
$CardSchemaVersion = 'content-summary-v2'

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

function Get-NodeAttribute($Node, [string]$Name) {{
  if ($null -eq $Node) {{ return '' }}
  $attr = $Node.Attributes.GetNamedItem($Name)
  if ($null -eq $attr) {{ return '' }}
  return [string]$attr.Value
}}

function Get-SectionPath($Section) {{
  $parts = New-Object System.Collections.ArrayList
  $node = $Section
  while ($null -ne $node -and $node.LocalName -ne 'Notebook') {{
    if (($node.LocalName -eq 'Section' -or $node.LocalName -eq 'SectionGroup') -and
        -not [string]::IsNullOrWhiteSpace($node.name)) {{
      [void]$parts.Insert(0, [string]$node.name)
    }}
    $node = $node.ParentNode
  }}
  return ($parts -join ' / ')
}}

function Should-SkipSourceSection($Section) {{
  $sectionPath = Get-SectionPath $Section
  $sectionName = [string]$Section.name
  if ($sectionName -like '分類_*') {{ return $true }}
  if ($sectionName -eq '削除されたページ') {{ return $true }}
  if ($sectionName -eq '実験') {{ return $true }}
  if ($sectionPath -match '(^| / )実験($| / )') {{ return $true }}
  return $false
}}

function Get-SectionPages($OneNote, $Section) {{
  try {{
    [xml]$sectionHierarchy = ''
    $OneNote.GetHierarchy($Section.ID, 4, [ref]$sectionHierarchy)
    return @($sectionHierarchy.DocumentElement.SelectNodes('.//*[local-name()="Page"]'))
  }} catch {{
    return @($Section.SelectNodes('.//*[local-name()="Page"]'))
  }}
}}

function Get-PageContentXml($OneNote, [string]$PageId) {{
  $pageXml = ''
  try {{
    $OneNote.GetPageContent($PageId, [ref]$pageXml, 0)
    return [xml]$pageXml
  }} catch {{
    try {{
      $OneNote.GetPageContent($PageId, [ref]$pageXml, 0, 7)
      return [xml]$pageXml
    }} catch {{
      return $null
    }}
  }}
}}

function Get-PagePlainText($OneNote, $Page) {{
  $content = Get-PageContentXml $OneNote $Page.ID
  if ($null -eq $content) {{ return '' }}
  $lines = New-Object System.Collections.ArrayList
  foreach ($node in @($content.SelectNodes('//*[local-name()="T"]'))) {{
    $text = ([string]$node.InnerText -replace '\\s+', ' ').Trim()
    if (-not [string]::IsNullOrWhiteSpace($text)) {{
      [void]$lines.Add($text)
    }}
  }}
  return ($lines -join "`n")
}}

function Get-EffectiveTitle($Page, [string]$PlainText) {{
  $title = [string]$Page.name
  if (-not [string]::IsNullOrWhiteSpace($title) -and $title -ne '無題のページ') {{
    return $title
  }}
  foreach ($line in @($PlainText -split "`n")) {{
    $line = ($line -replace '\\s+', ' ').Trim()
    if (-not [string]::IsNullOrWhiteSpace($line)) {{
      if ($line.Length -gt 60) {{ return $line.Substring(0, 60) + '...' }}
      return $line
    }}
  }}
  return '無題のページ'
}}

function Get-ContentPreview([string]$PlainText) {{
  $normalized = ($PlainText -replace '\\s+', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($normalized)) {{
    return '本文テキストを抽出できませんでした（空ページ、画像/手書きのみ、または未同期の可能性があります）。'
  }}
  if ($normalized.Length -gt 900) {{ return $normalized.Substring(0, 900) + '...' }}
  return $normalized
}}

function Get-SimpleSummary([string]$Title, [string]$SourceSectionName, [string]$PlainText) {{
  $normalized = ($PlainText -replace '\\s+', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($normalized)) {{
    return "本文テキストなし。元セクション「$SourceSectionName」のページ「$Title」です。"
  }}
  $summary = $normalized
  if ($summary.Length -gt 180) {{ $summary = $summary.Substring(0, 180) + '...' }}
  if ([string]::IsNullOrWhiteSpace($Title) -or $Title -eq '無題のページ') {{
    return $summary
  }}
  return "${{Title}}: $summary"
}}

function Get-TextHash([string]$Text) {{
  if ($null -eq $Text) {{ $Text = '' }}
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {{
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
  }} finally {{
    $sha.Dispose()
  }}
}}

function Get-PageSignature($Page, [string]$SourceSectionName, [string]$PlainText) {{
  $parts = @(
    $SourceSectionName,
    [string]$Page.name,
    (Get-NodeAttribute $Page 'lastModifiedTime'),
    (Get-NodeAttribute $Page 'dateTime'),
    (Get-TextHash $PlainText)
  )
  return ($parts -join '|')
}}

function Select-Category($Page, [string]$SourceSectionName, [string]$PlainText) {{
  $target = (Normalize-Text ($SourceSectionName + ' ' + $Page.name + ' ' + $PlainText))
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
    $OneNote.GetHyperlinkToObject($SourcePage.SourcePageId, '', [ref]$link)
  }} catch {{
    $link = $SourcePage.SourcePageId
  }}

  $title = $SourcePage.Title
  if ([string]::IsNullOrWhiteSpace($title)) {{ $title = '無題のページ' }}
  $lines = @(
    "要約: $($SourcePage.Summary)",
    "分類: $CategorySectionName",
    "元セクション: $SourceSectionName",
    "元ページ: $title",
    "元ページリンク: $link",
    "本文抜粋: $($SourcePage.ContentPreview)",
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
  -not (Should-SkipSourceSection $_)
}})

$existingTargetPageIds = @{{}}
$existingClassificationCards = 0
foreach ($category in $Categories) {{
  $classificationSection = $notebook.SelectNodes('.//*[local-name()="Section"]') |
    Where-Object {{ $_.name -eq $category.Section }} |
    Select-Object -First 1
  if ($null -ne $classificationSection) {{
    foreach ($page in @(Get-SectionPages $one $classificationSection)) {{
      if (-not [string]::IsNullOrWhiteSpace($page.ID)) {{
        $existingTargetPageIds[$page.ID] = $true
        $existingClassificationCards += 1
      }}
    }}
  }}
}}

$allPlan = New-Object System.Collections.ArrayList
$seenPageIds = @{{}}
$emptySections = New-Object System.Collections.ArrayList
foreach ($section in $sourceSections) {{
  $sectionPath = Get-SectionPath $section
  $sectionPages = @(Get-SectionPages $one $section)
  if ($sectionPages.Count -eq 0) {{
    [void]$emptySections.Add($sectionPath)
  }}
  foreach ($page in $sectionPages) {{
    if ([string]::IsNullOrWhiteSpace($page.ID) -or $seenPageIds.ContainsKey($page.ID)) {{
      continue
    }}
    $seenPageIds[$page.ID] = $true
    $plainText = Get-PagePlainText $one $page
    $effectiveTitle = Get-EffectiveTitle $page $plainText
    $contentPreview = Get-ContentPreview $plainText
    $summaryText = Get-SimpleSummary $effectiveTitle $sectionPath $plainText
    $category = Select-Category $page $sectionPath $plainText
    [void]$allPlan.Add([pscustomobject]@{{
      SourceSection = $sectionPath
      SourcePageId = $page.ID
      SourceLastModified = (Get-NodeAttribute $page 'lastModifiedTime')
      SourceSignature = ((Get-PageSignature $page $sectionPath $plainText) + '|' + $category.Key + '|' + $CardSchemaVersion)
      Title = $effectiveTitle
      Summary = $summaryText
      ContentPreview = $contentPreview
      CategoryKey = $category.Key
      CategorySection = $category.Section
    }})
  }}
}}

$previousEntriesById = @{{}}
if ($Incremental -and (Test-Path -LiteralPath $MapPath)) {{
  try {{
    $previousMap = Get-Content -LiteralPath $MapPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($previousMap.notebook -eq $NotebookName) {{
      foreach ($entry in @($previousMap.entries)) {{
        if ($entry.sourcePageId) {{ $previousEntriesById[$entry.sourcePageId] = $entry }}
      }}
    }}
  }} catch {{
    "WARN: 既存の分類履歴を読めませんでした。全件を差分対象として扱います: $($_.Exception.Message)"
  }}
}}

$currentIds = @{{}}
$plan = New-Object System.Collections.ArrayList
$unchangedEntries = New-Object System.Collections.ArrayList
foreach ($item in $allPlan) {{
  $currentIds[$item.SourcePageId] = $true
  $previous = $previousEntriesById[$item.SourcePageId]
  $previousSignature = $null
  if ($null -ne $previous) {{
    $previousSignature = $previous.sourceSignature
    if ([string]::IsNullOrWhiteSpace($previousSignature)) {{
      $previousSignature = (($previous.sourceSection, $previous.title, $previous.sourceLastModified, '') -join '|')
    }}
  }}
  $previousTargetExists = (
    $null -ne $previous -and
    -not [string]::IsNullOrWhiteSpace($previous.targetPageId) -and
    $existingTargetPageIds.ContainsKey($previous.targetPageId)
  )

  if ((-not $Incremental) -or ($null -eq $previous) -or (-not $previousTargetExists) -or ($previousSignature -ne $item.SourceSignature)) {{
    $previousTargetPageId = ''
    if ($null -ne $previous) {{ $previousTargetPageId = $previous.targetPageId }}
    $item | Add-Member -NotePropertyName PreviousTargetPageId -NotePropertyValue $previousTargetPageId
    [void]$plan.Add($item)
  }} else {{
    [void]$unchangedEntries.Add($previous)
  }}
}}

$removedEntries = @($previousEntriesById.Values | Where-Object {{ -not $currentIds.ContainsKey($_.sourcePageId) }})

$summary = @{{}}
foreach ($category in $Categories) {{ $summary[$category.Section] = 0 }}
foreach ($item in $allPlan) {{ $summary[$item.CategorySection] += 1 }}

"Notebook: $NotebookName"
"対象元セクション: $($sourceSections.Count)"
"対象ページ: $($allPlan.Count)"
"差分分類対象: $($plan.Count)"
"履歴から消えた元ページ: $($removedEntries.Count)"
"ページ0件の元セクション: $($emptySections.Count)"
"既存分類カード: $existingClassificationCards"
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
  $unchangedEntries = New-Object System.Collections.ArrayList
  foreach ($category in $Categories) {{
    foreach ($page in @(Get-SectionPages $one $targetSections[$category.Key])) {{
      $one.DeleteHierarchy($page.ID)
    }}
  }}
}}

$failed = New-Object System.Collections.ArrayList
$deletedStaleCards = 0
if (-not $ResetClassificationPages) {{
  $keepTargetIds = @{{}}
  foreach ($entry in @($unchangedEntries)) {{
    if (-not [string]::IsNullOrWhiteSpace($entry.targetPageId)) {{
      $keepTargetIds[$entry.targetPageId] = $true
    }}
  }}
  foreach ($category in $Categories) {{
    foreach ($page in @(Get-SectionPages $one $targetSections[$category.Key])) {{
      if (-not $keepTargetIds.ContainsKey($page.ID)) {{
        try {{
          $one.DeleteHierarchy($page.ID)
          $deletedStaleCards += 1
        }} catch {{
          [void]$failed.Add([pscustomobject]@{{
            sourceSection = $category.Section
            title = $page.name
            category = $category.Key
            error = "古い分類カードの削除に失敗: $($_.Exception.Message)"
          }})
        }}
      }}
    }}
  }}
}}

$created = @{{}}
foreach ($category in $Categories) {{ $created[$category.Key] = 0 }}
$entries = New-Object System.Collections.ArrayList

foreach ($item in $plan) {{
  try {{
    if (-not [string]::IsNullOrWhiteSpace($item.PreviousTargetPageId)) {{
      try {{ $one.DeleteHierarchy($item.PreviousTargetPageId) }} catch {{ }}
    }}
    $targetId = New-ClassificationCard $one $categorySectionIds[$item.CategoryKey] $item $item.SourceSection $item.CategorySection
    [void]$entries.Add([pscustomobject]@{{
      sourcePageId = $item.SourcePageId
      sourceSection = $item.SourceSection
      sourceLastModified = $item.SourceLastModified
      sourceSignature = $item.SourceSignature
      title = $item.Title
      summary = $item.Summary
      contentPreview = $item.ContentPreview
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

$allEntries = @($unchangedEntries) + @($entries)
$map = [pscustomobject]@{{
  notebook = $NotebookName
  updatedAt = (Get-Date).ToString('o')
  mode = $(if ($Incremental) {{ 'incremental' }} else {{ 'full' }})
  entries = @($allEntries)
}}
$mapDir = Split-Path -Parent $MapPath
if ($mapDir -and -not (Test-Path -LiteralPath $mapDir)) {{
  New-Item -ItemType Directory -Path $mapDir | Out-Null
}}
$map | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $MapPath -Encoding UTF8

"実行結果:"
"古い/重複した分類カード削除: $deletedStaleCards"
foreach ($category in $Categories) {{
  "$($category.Section): 作成 $($created[$category.Key])"
}}
"失敗: $($failed.Count)"
foreach ($item in $failed) {{
  "FAILED: [$($item.sourceSection)] $($item.title) -> $($item.category): $($item.error)"
}}
"""


def safe_filename_part(text: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', '_', text).strip().strip('.')
    return value or 'notebook'


def normalize_notebook_name(text: str) -> str:
    value = unicodedata.normalize("NFKC", text)
    value = re.sub(r"[（(].*?[）)]", "", value)
    value = value.replace("書き込み", "書込").replace("書込み", "書込")
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"[-・･_＿\\/／（）()「」『』\\[\\]【】]", "", value)
    return value.casefold()


def list_onenote_notebooks() -> list[str]:
    script = (
        "$ErrorActionPreference = 'Stop';"
        "$one = New-Object -ComObject OneNote.Application;"
        "[xml]$hierarchy = '';"
        "$one.GetHierarchy('', 4, [ref]$hierarchy);"
        "$hierarchy.DocumentElement.SelectNodes('//*') | "
        "Where-Object { $_.LocalName -eq 'Notebook' } | "
        "ForEach-Object { $_.name }"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode != 0:
        return []
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def choose_notebook_from_candidates(requested: str, notebooks: list[str]) -> str | None:
    if not notebooks:
        return None

    for notebook in notebooks:
        if notebook == requested:
            return notebook

    normalized_requested = normalize_notebook_name(requested)
    normalized_by_name = {notebook: normalize_notebook_name(notebook) for notebook in notebooks}

    exact_normalized = [
        notebook
        for notebook, normalized in normalized_by_name.items()
        if normalized == normalized_requested
    ]
    if len(exact_normalized) == 1:
        return exact_normalized[0]

    contains_matches = [
        notebook
        for notebook, normalized in normalized_by_name.items()
        if normalized_requested
        and (normalized_requested in normalized or normalized in normalized_requested)
    ]
    if len(contains_matches) == 1:
        return contains_matches[0]

    close_normalized = difflib.get_close_matches(
        normalized_requested,
        list(normalized_by_name.values()),
        n=5,
        cutoff=0.55,
    )
    close_matches = [
        notebook
        for notebook, normalized in normalized_by_name.items()
        if normalized in close_normalized
    ]
    candidates = exact_normalized or contains_matches or close_matches
    if not candidates:
        return None
    if len(candidates) == 1:
        answer = input(f"'{requested}' に近いノート '{candidates[0]}' で分類しますか？ [Y/n]: ").strip()
        return candidates[0] if answer.casefold() not in {"n", "no"} else None

    print(f"'{requested}' に近いノート候補が複数あります:")
    for index, candidate in enumerate(candidates, start=1):
        print(f"  {index}. {candidate}")
    answer = input("分類するノート番号を入力してください（空欄で中止）: ").strip()
    if not answer:
        return None
    try:
        selected = int(answer)
    except ValueError:
        return None
    if 1 <= selected <= len(candidates):
        return candidates[selected - 1]
    return None


def resolve_notebook(name: str | None) -> str:
    requested = name or input("どのノートを分類しますか？: ").strip()
    if not requested:
        raise SystemExit("ノート名が入力されていません。")

    notebooks = list_onenote_notebooks()
    resolved = choose_notebook_from_candidates(requested, notebooks)
    if resolved:
        if resolved != requested:
            print(f"ノート名を解決しました: {requested} -> {resolved}")
        return resolved

    if notebooks:
        print("開いているOneNoteノートブック:")
        for notebook in notebooks:
            print(f"  - {notebook}")
    raise SystemExit(f"ノートブックが見つかりません: {requested}")


def resolve_map_file(map_file: Path | None, notebook: str) -> Path:
    if map_file is not None:
        return map_file
    return DEFAULT_MAP_DIR / f"{safe_filename_part(notebook)}.json"

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
    parser.add_argument("--notebook", help="分類対象のOneNoteノートブック名。未指定時は実行時に質問します")
    parser.add_argument("--execute", action="store_true", help="OneNoteへ分類セクション/分類カードを書き込む")
    parser.add_argument("--full-rebuild", action="store_true", help="履歴を使わず全件を分類し、分類セクション内ページを作り直す")
    parser.add_argument(
        "--keep-existing",
        action="store_true",
        help="全件再分類時に既存の分類セクション内ページを消さずに追記する",
    )
    parser.add_argument("--map-file", type=Path, help="分類結果JSONの保存先。未指定時はノートごとに tools/onenote_classification_maps/<ノート名>.json へ保存")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    notebook = resolve_notebook(args.notebook)
    map_file = resolve_map_file(args.map_file, notebook)
    script = build_powershell_script(
        notebook=notebook,
        execute=args.execute,
        reset_classification_pages=args.full_rebuild and not args.keep_existing,
        map_file=map_file,
        incremental=not args.full_rebuild,
    )
    result = run_powershell(script)
    if result.stdout:
        print(result.stdout.rstrip())
    if result.stderr:
        print(result.stderr.rstrip(), file=sys.stderr)
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
