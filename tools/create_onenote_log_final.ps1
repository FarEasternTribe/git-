param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Continue'

try {
  Write-Host "Starting OneNote log creation..."

  # テキスト読み込み
  $textContent = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)

  # OneNote COM オブジェクト
  $one = New-Object -ComObject OneNote.Application

  # セクション取得
  $hierarchyText = [ref]''
  $one.GetHierarchy('', 4, $hierarchyText, 2)
  $hierarchyText = $hierarchyText.Value

  [xml]$hierarchy = $hierarchyText
  $sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })

  if ($sections.Count -ne 1) {
    throw "Section not found: $SectionName"
  }

  $section = $sections[0]

  # ページ作成
  $pageIdRef = [ref]''
  $one.CreateNewPage($section.ID, $pageIdRef, 1)
  $pageId = $pageIdRef.Value

  if ([string]::IsNullOrWhiteSpace($pageId)) {
    throw "Failed to create page"
  }

  Write-Host "Page created: $pageId"

  # ページコンテンツ作成（OneNote XML スキーマに従う）
  $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

  $pageContent = '<?xml version="1.0" encoding="UTF-8"?>'
  $pageContent += "<one:Page xmlns:one='$ns' ID='$pageId'>"
  $pageContent += "<one:Title>"
  $pageContent += "<one:OE>"
  $pageContent += "<one:T><![CDATA[$PageTitle]]></one:T>"
  $pageContent += "</one:OE>"
  $pageContent += "</one:Title>"
  $pageContent += "<one:Outline>"
  $pageContent += "<one:OEChildren>"

  # テキスト行を追加
  $lines = $textContent -split "`n"
  foreach ($line in $lines) {
    if ($line.Trim() -ne '') {
      # 特殊文字をエスケープ
      $line = $line -replace '&', '&amp;'
      $line = $line -replace '<', '&lt;'
      $line = $line -replace '>', '&gt;'
      $pageContent += "<one:OE><one:T><![CDATA[$line]]></one:T></one:OE>"
    }
  }

  $pageContent += "</one:OEChildren>"
  $pageContent += "</one:Outline>"
  $pageContent += "</one:Page>"

  # OneNote に更新
  Write-Host "Updating page content..."
  $one.UpdatePageContent($pageContent)

  Write-Host "SUCCESS: Log recorded to OneNote"
  Write-Host "PageId=$pageId"
  Write-Host "Section=$SectionName"
}
catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  exit 1
}
