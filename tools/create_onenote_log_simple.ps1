param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Stop'

# テキスト読み込み
$textContent = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)

# OneNote COM オブジェクト
$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)

# セクション取得
[xml]$hierarchy = $hierarchyText
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })

if ($sections.Count -ne 1) {
  throw "Section not found: $SectionName"
}

$section = $sections[0]

# ページ作成
$pageId = ''
$one.CreateNewPage($section.ID, [ref]$pageId, 1)

if ([string]::IsNullOrWhiteSpace($pageId)) {
  throw "Failed to create page"
}

# ページコンテンツ作成
$ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

# テキスト行ごとに OE 要素を作成
$pageContent = '<?xml version="1.0" encoding="UTF-8"?>'
$pageContent += "<one:Page xmlns:one='$ns' formatVersion='2.0' ID='$pageId'>"
$pageContent += "<one:Title><one:OEChildren><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:OEChildren></one:Title>"
$pageContent += "<one:Outline><one:OEChildren>"

# 各行を追加
$lines = $textContent -split "`n"
foreach ($line in $lines) {
  if ($line.Trim() -ne '') {
    $pageContent += "<one:OE><one:T><![CDATA[$line]]></one:T></one:OE>"
  }
}

$pageContent += "</one:OEChildren></one:Outline>"
$pageContent += "</one:Page>"

# OneNote に更新
$one.UpdatePageContent($pageContent)

Write-Host "SUCCESS"
Write-Host "PageId=$pageId"
Write-Host "Section=$SectionName"
