$ErrorActionPreference = 'Stop'

$NotebookName = '2025年書込テスト'

$Categories = @(
  @{ Key = 'inbox'; Section = '分類_01_未整理・一時置き場' },
  @{ Key = 'research'; Section = '分類_02_研究・論文・技術メモ' },
  @{ Key = 'work'; Section = '分類_03_仕事・大学業務' },
  @{ Key = 'writing'; Section = '分類_04_文章・メール・発表テンプレ' },
  @{ Key = 'journal'; Section = '分類_05_日誌・思考整理・ToDo' },
  @{ Key = 'life'; Section = '分類_06_生活・個人メモ' },
  @{ Key = 'web'; Section = '分類_07_Webクリップ・あとで読む' }
)

function Normalize-Title([string]$Text) {
  if ($null -eq $Text) { return '' }
  return ($Text -replace '\s+', ' ').Trim().ToLowerInvariant()
}

function Get-CategoryKey([string]$Title, [string]$SourceSection) {
  $t = Normalize-Title $Title
  $s = Normalize-Title $SourceSection

  if ($s -eq '料理' -or $t -match '料理|レシピ|delishkitchen|オランデーズ|麻醤|楽天|トナー|車|フロントガラス|慢性疲労|食べたい|生活|家事|買い回り') {
    return 'life'
  }
  if ($s -eq 'timeline' -or $s -like 'memoまとめ*' -or $t -match 'timeline|タイムライン|todo|to-do|対話まとめ|日誌|ジャーナリング|問題発生時チェックシート|思考|要約サマリー|最新ルール') {
    return 'journal'
  }
  if ($s -eq 'メールテンプレ' -or $t -match 'メール|テンプレ|講演者各位|members,|hello, my name|プレゼントーク|slide|タイトル|シラバス|お世話になっております|自己紹介') {
    return 'writing'
  }
  if ($t -match 'r7年度|センター談話会|作業環境測定|松田先生|中嶋先生|成果|指標|京都大学|坂口研究室|授業|講習|依頼|業務|文科省|就学支援金|申請') {
    return 'work'
  }
  if ($s -match '合成|co2|新しいセクション 1|新しいセクション 6' -or $t -match 'gnr|グラフェン|ジメトキシ|br化|qd|量子ドット|gese|強誘電|反強誘電|shift|シフト電流|合成|特許|patent|espacenet|600 mesh|tem|cu tem|らせん|ferromagnet|macromolecular|exfoliation|研究概要|論文|acs|omega|化学|半導体|切削') {
    return 'research'
  }
  if ($t -match '^https?://|note\.com|smartnews|review\.kakaku|amazon\.co\.jp|lifehacker|ライフハッカー|delishkitchen|masterorganicchemistry|sigmaaldrich|youtube|dj shadow|コンテンツへとスキップ|人気の記事一覧|スキ一覧') {
    return 'web'
  }
  return 'inbox'
}

function Get-NodeByName($Root, [string]$LocalName, [string]$Name) {
  return $Root.SelectNodes('//*') | Where-Object { $_.LocalName -eq $LocalName -and $_.name -eq $Name } | Select-Object -First 1
}

function Refresh-Hierarchy($OneNote) {
  [xml]$hierarchy = ''
  $OneNote.GetHierarchy('', 4, [ref]$hierarchy)
  return $hierarchy
}

function Ensure-Section($OneNote, [xml]$Hierarchy, $Notebook, [string]$SectionName) {
  $existing = $Notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName } | Select-Object -First 1
  if ($null -ne $existing) {
    return $existing.ID
  }

  $ns = $Notebook.NamespaceURI
  $section = $Hierarchy.CreateElement('one', 'Section', $ns)
  $section.SetAttribute('name', $SectionName)
  $section.SetAttribute('path', $Notebook.path + $SectionName + '.one')

  $firstSectionGroup = $Notebook.SelectNodes('./*[local-name()="SectionGroup"]') | Select-Object -First 1
  if ($null -ne $firstSectionGroup) {
    [void]$Notebook.InsertBefore($section, $firstSectionGroup)
  } else {
    [void]$Notebook.AppendChild($section)
  }

  $OneNote.UpdateHierarchy($Notebook.OuterXml)
  Start-Sleep -Milliseconds 800

  $updatedHierarchy = Refresh-Hierarchy $OneNote
  $updatedNotebook = Get-NodeByName $updatedHierarchy.DocumentElement 'Notebook' $NotebookName
  $created = $updatedNotebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName } | Select-Object -First 1
  if ($null -eq $created) {
    throw "Failed to create section: $SectionName"
  }
  return $created.ID
}

function Escape-CData([string]$Text) {
  if ($null -eq $Text) { return '' }
  return $Text.Replace(']]>', ']]]]><![CDATA[>')
}

function New-PageXml([string]$Title, [string[]]$Lines) {
  $oeParts = @()
  foreach ($line in $Lines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $oeParts += @"
      <one:OE>
        <one:T><![CDATA[$(Escape-CData $line)]]></one:T>
      </one:OE>
"@
  }
  if ($oeParts.Count -eq 0) {
    $oeParts += @"
      <one:OE>
        <one:T><![CDATA[(empty note)]]></one:T>
      </one:OE>
"@
  }

  $body = $oeParts -join "`n"
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
}

function New-ClassificationPage($OneNote, [string]$TargetSectionId, $SourcePage, [string]$SourceSectionName, [string]$CategorySectionName) {
  $newPageId = ''
  $OneNote.CreateNewPage($TargetSectionId, [ref]$newPageId, 0)

  $link = ''
  try {
    $OneNote.GetHyperlinkToObject($SourcePage.ID, '', [ref]$link)
  } catch {
    $link = $SourcePage.ID
  }

  $title = $SourcePage.name
  if ([string]::IsNullOrWhiteSpace($title)) {
    $title = '無題のページ'
  }

  $lines = @(
    "分類: $CategorySectionName",
    "元セクション: $SourceSectionName",
    "元ページ: $title",
    "元ページリンク: $link",
    "注記: 元ノートは削除・移動していません。このページは分類用の参照カードです。"
  )
  $pageXml = (New-PageXml $title $lines).Replace('__PAGE_ID__', $newPageId)
  $OneNote.UpdatePageContent($pageXml)
  return $newPageId
}

$one = New-Object -ComObject OneNote.Application
$hierarchy = Refresh-Hierarchy $one
$notebook = Get-NodeByName $hierarchy.DocumentElement 'Notebook' $NotebookName
if ($null -eq $notebook) {
  throw "Notebook not found: $NotebookName"
}

$mapPath = Join-Path (Get-Location) 'tools\onenote_2025_classification_map.json'
$copiedSourceIds = @{}
$newMapEntries = New-Object System.Collections.ArrayList

$categorySectionIds = @{}
foreach ($category in $Categories) {
  $categorySectionIds[$category.Key] = Ensure-Section $one $hierarchy $notebook $category.Section
}

$hierarchy = Refresh-Hierarchy $one
$notebook = Get-NodeByName $hierarchy.DocumentElement 'Notebook' $NotebookName
$targetSections = @{}
foreach ($category in $Categories) {
  $section = $notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $category.Section } | Select-Object -First 1
  $targetSections[$category.Key] = $section
}

foreach ($category in $Categories) {
  $section = $targetSections[$category.Key]
  foreach ($page in @($section.SelectNodes('.//*[local-name()="Page"]'))) {
    $One.DeleteHierarchy($page.ID)
  }
}

$hierarchy = Refresh-Hierarchy $one
$notebook = Get-NodeByName $hierarchy.DocumentElement 'Notebook' $NotebookName
foreach ($category in $Categories) {
  $section = $notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $category.Section } | Select-Object -First 1
  $targetSections[$category.Key] = $section
}

$sourceSections = $notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
  $name = $_.name
  ($name -notlike '分類_*') -and ($name -ne '削除されたページ')
}

$created = @{}
$skipped = @{}
$failed = New-Object System.Collections.ArrayList
foreach ($category in $Categories) {
  $created[$category.Key] = 0
  $skipped[$category.Key] = 0
}

foreach ($section in $sourceSections) {
  foreach ($page in $section.SelectNodes('.//*[local-name()="Page"]')) {
    $categoryKey = Get-CategoryKey $page.name $section.name
    if ($copiedSourceIds.ContainsKey($page.ID)) {
      $skipped[$categoryKey] += 1
      continue
    }

    try {
      $categorySectionName = ($Categories | Where-Object { $_.Key -eq $categoryKey } | Select-Object -First 1).Section
      $targetPageId = New-ClassificationPage $one $categorySectionIds[$categoryKey] $page $section.name $categorySectionName
      $entry = [pscustomobject]@{
        sourcePageId = $page.ID
        sourceSection = $section.name
        title = $page.name
        category = $categoryKey
        targetPageId = $targetPageId
        copiedAt = (Get-Date).ToString('o')
      }
      $copiedSourceIds[$page.ID] = $entry
      [void]$newMapEntries.Add($entry)
      $created[$categoryKey] += 1
    } catch {
      [void]$failed.Add([pscustomobject]@{
        SourceSection = $section.name
        Page = $page.name
        Category = $categoryKey
        Error = $_.Exception.Message
      })
    }
  }
}

$allEntries = @($copiedSourceIds.Values)
$mapOutput = [pscustomobject]@{
  notebook = $NotebookName
  updatedAt = (Get-Date).ToString('o')
  entries = $allEntries
}
$mapOutput | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $mapPath -Encoding UTF8

'Notebook: ' + $NotebookName
'分類セクション: ' + $Categories.Count
foreach ($category in $Categories) {
  '{0}: 新規コピー {1} / 既存スキップ {2}' -f $category.Section, $created[$category.Key], $skipped[$category.Key]
}
'失敗: ' + $failed.Count
foreach ($item in $failed) {
  'FAILED: [{0}] {1} -> {2}: {3}' -f $item.SourceSection, $item.Page, $item.Category, $item.Error
}
