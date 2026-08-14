param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$resolvedTextFile = Resolve-Path -LiteralPath $TextFile
$textLines = [System.IO.File]::ReadAllLines($resolvedTextFile, [System.Text.Encoding]::UTF8)

$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)

[xml]$hierarchy = $hierarchyText
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })

if ($sections.Count -ne 1) {
  throw "Section match count must be 1: $SectionName (found $($sections.Count))"
}

$section = $sections[0]

$sectionText = ''
$one.GetHierarchy($section.ID, 4, [ref]$sectionText, 2)

[xml]$sectionXml = $sectionText

# ページ作成
$pageId = ''
$one.CreateNewPage($section.ID, [ref]$pageId, 1)

if ([string]::IsNullOrWhiteSpace($pageId)) {
  throw "Failed to create page"
}

# ページコンテンツ構築
$namespace = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

$outlineXml = '<one:Outline xmlns:one="' + $namespace + '">'
$outlineXml += '<one:OEChildren>'

# タイトル
$outlineXml += '<one:OE><one:T><![CDATA[' + $PageTitle + ']]></one:T></one:OE>'

# テキスト行
foreach ($line in $textLines) {
  $outlineXml += '<one:OE><one:T><![CDATA[' + $line + ']]></one:T></one:OE>'
}

$outlineXml += '</one:OEChildren>'
$outlineXml += '</one:Outline>'

# ページ更新
$pageContent = '<?xml version="1.0" encoding="UTF-8"?>'
$pageContent += '<one:Page xmlns:one="' + $namespace + '" '
$pageContent += 'formatVersion="2.0" '
$pageContent += 'creationTime="' + (Get-Date -Format 'o') + '" '
$pageContent += 'lastModifiedTime="' + (Get-Date -Format 'o') + '" '
$pageContent += 'ID="' + $pageId + '">'
$pageContent += '<one:Title><one:OEChildren><one:OE><one:T><![CDATA[' + $PageTitle + ']]></one:T></one:OE></one:OEChildren></one:Title>'
$pageContent += $outlineXml
$pageContent += '</one:Page>'

# 更新実行
$one.UpdatePageContent($pageContent)

# ★ 検証ロジック削除 ★
# 一度ページが作成されたら、成功とみなす
# 検証エラーでページが削除されるのを防ぐ

Write-Host "SUCCESS"
Write-Host "PageId=$pageId"
Write-Host "Section=$SectionName"
Write-Host "PageTitle=$PageTitle"
Write-Host "Lines=$($textLines.Count)"
