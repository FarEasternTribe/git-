param(
  [Parameter(Mandatory = $true)]
  [string]$MarkdownPath,
  [string]$NotebookName = 'OpenAI_Agent1',
  [string]$SectionName = '日誌',
  # Claude(Cowork)が文脈理解で生成した日誌であることを示すマーカー。
  # 実機のローカル抽出版(マーカー無し)や Codex版(_codex)と区別するために付ける。
  [string]$Marker = '[Claude]'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $MarkdownPath)) {
  throw "Markdown not found: $MarkdownPath"
}

$markdown = Get-Content -LiteralPath $MarkdownPath -Raw -Encoding UTF8

# タイトルは既存慣習 "YYYY-MM-DD_日誌" にマーカーを前置する（例: "[Claude] 2026-07-14_日誌"）。
# ファイル名(拡張子除く)をベースにする。
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($MarkdownPath)
$title = "$Marker $baseName"

$now = Get-Date
# ノート冒頭にも [Claude] 生成であることを明記する行を入れる（"ノート部分に記載"の要件）。
$headerLine = "$Marker Cowork(Claude)による文脈理解版の日誌 / 生成: $($now.ToString('yyyy-MM-dd HH:mm'))"
$body = $headerLine + "`n`n" + $markdown

try {
  $one = New-Object -ComObject OneNote.Application
  [xml]$hierarchy = ''
  $one.GetHierarchy('', 4, [ref]$hierarchy)

  $notebook = @($hierarchy.DocumentElement.SelectNodes('//*[local-name()="Notebook"]') | Where-Object {
    $_.name -eq $NotebookName
  } | Select-Object -First 1)
  if ($null -eq $notebook) {
    throw "Notebook not found: $NotebookName"
  }

  $section = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
    $_.name -eq $SectionName
  } | Select-Object -First 1)
  if ($null -eq $section) {
    throw "Section not found: $NotebookName / $SectionName"
  }

  $pageId = ''
  $one.CreateNewPage($section.ID, [ref]$pageId, 0)

  $bodyLines = @()
  foreach ($line in $body -split "`n") {
    $bodyLines += '<one:OE><one:T><![CDATA[' + $line + ']]></one:T></one:OE>'
  }
  $bodyXml = $bodyLines -join "`n"
  $pageXml = @"
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="$pageId">
  <one:Title>
    <one:OE>
      <one:T><![CDATA[$title]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0" />
    <one:Size width="900" height="600" />
    <one:OEChildren>
$bodyXml
    </one:OEChildren>
  </one:Outline>
</one:Page>
"@
  $one.UpdatePageContent($pageXml)
  Write-Host "OneNote: ok"
  Write-Host "Notebook: $NotebookName"
  Write-Host "Section: $SectionName"
  Write-Host "Title: $title"
  Write-Host "PageId: $pageId"
} catch {
  Write-Warning "OneNote journal write failed: $($_.Exception.Message)"
  Write-Host "OneNote: failed"
  throw
}
