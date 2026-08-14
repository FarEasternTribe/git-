param(
  [string]$NotebookName = 'OpenAI_Agent1',
  [string]$SectionName = 'Claude プロセスログ',
  [string]$MarkdownFile = ''
)

$ErrorActionPreference = 'Stop'

try {
  Write-Host "Starting OneNote section creation..."

  # マークダウンファイル読み込み
  if ([string]::IsNullOrWhiteSpace($MarkdownFile)) {
    $MarkdownFile = "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳\Claude プロセスログ\2026-08-05_Research_OS_Development_Log.md"
  }

  Write-Host "Reading markdown file: $MarkdownFile"
  $content = Get-Content -LiteralPath $MarkdownFile -Encoding UTF8 -Raw

  # OneNote COM オブジェクト作成
  Write-Host "Creating OneNote COM object..."
  $one = New-Object -ComObject OneNote.Application

  # ノートブック階層取得
  Write-Host "Getting notebook hierarchy..."
  $hierarchyRef = [ref]''
  $one.GetHierarchy('', 4, $hierarchyRef, 2)
  $hierarchy = $hierarchyRef.Value

  [xml]$hierarchyXml = $hierarchy
  $notebook = @($hierarchyXml.SelectNodes('//*[local-name()="Notebook"]') | Where-Object { $_.name -eq $NotebookName })

  if ($notebook.Count -ne 1) {
    throw "Notebook not found or multiple matches: $NotebookName"
  }

  $notebook = $notebook[0]
  Write-Host "Found notebook: $($notebook.name)"

  # セクション作成
  Write-Host "Creating section: $SectionName"
  $sectionIdRef = [ref]''
  $one.OpenHierarchy($notebook.path + $SectionName + '.one', '', $sectionIdRef, 3)
  $sectionId = $sectionIdRef.Value

  if ([string]::IsNullOrWhiteSpace($sectionId)) {
    throw "Failed to create section"
  }

  Write-Host "Section created: $sectionId"

  # ページ作成
  Write-Host "Creating page..."
  $pageIdRef = [ref]''
  $one.CreateNewPage($sectionId, $pageIdRef, 0)
  $pageId = $pageIdRef.Value

  if ([string]::IsNullOrWhiteSpace($pageId)) {
    throw "Failed to create page"
  }

  Write-Host "Page created: $pageId"

  # ページコンテンツ作成
  $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

  # マークダウンの行をOE要素に変換
  $lines = $content -split "`n"
  $bodyLines = @()
  foreach ($line in $lines) {
    # 特殊文字をエスケープ
    $line = $line -replace '&', '&amp;'
    $line = $line -replace '<', '&lt;'
    $line = $line -replace '>', '&gt;'
    $bodyLines += "<one:OE><one:T><![CDATA[$line]]></one:T></one:OE>"
  }

  $bodyXml = $bodyLines -join "`n"

  # ページXML構築
  $pageXml = '<?xml version="1.0" encoding="UTF-8"?>'
  $pageXml += "<one:Page xmlns:one='$ns' ID='$pageId'>"
  $pageXml += '<one:Title>'
  $pageXml += '<one:OE>'
  $pageXml += '<one:T><![CDATA[2026-08-05 Research OS Dashboard 開発プロセスログ]]></one:T>'
  $pageXml += '</one:OE>'
  $pageXml += '</one:Title>'
  $pageXml += '<one:Outline>'
  $pageXml += '<one:Position x="36" y="86" z="0" />'
  $pageXml += '<one:Size width="900" height="600" />'
  $pageXml += '<one:OEChildren>'
  $pageXml += $bodyXml
  $pageXml += '</one:OEChildren>'
  $pageXml += '</one:Outline>'
  $pageXml += '</one:Page>'

  # OneNote に更新
  Write-Host "Updating page content..."
  $one.UpdatePageContent($pageXml)

  Write-Host "SUCCESS: Page created and updated"
  Write-Host "PageId=$pageId"
  Write-Host "Section=$SectionName"
  Write-Host "Notebook=$NotebookName"
}
catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  exit 1
}
