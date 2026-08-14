param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Continue'

try {
  # テキスト読み込み
  Write-Host "Reading text file: $TextFile"
  $textContent = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)
  Write-Host "Text read: $($textContent.Length) characters"

  # OneNote COM オブジェクト作成
  Write-Host "Creating OneNote COM object..."
  $one = $null
  $one = New-Object -ComObject OneNote.Application

  if ($one -eq $null) {
    throw "Failed to create OneNote.Application COM object"
  }

  Write-Host "OneNote COM object created successfully"

  # セクション取得
  Write-Host "Getting hierarchy..."
  $hierarchyText = ''
  $one.GetHierarchy('', 4, [ref]$hierarchyText, 2)

  if ([string]::IsNullOrWhiteSpace($hierarchyText)) {
    throw "GetHierarchy returned empty string"
  }

  Write-Host "Hierarchy obtained: $($hierarchyText.Length) characters"

  # XML パース
  [xml]$hierarchy = $hierarchyText
  $sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })

  Write-Host "Sections found: $($sections.Count)"

  if ($sections.Count -ne 1) {
    throw "Section match count must be 1: $SectionName (found $($sections.Count))"
  }

  $section = $sections[0]
  Write-Host "Section ID: $($section.ID)"

  # ページ作成
  Write-Host "Creating new page..."
  $pageId = ''
  $one.CreateNewPage($section.ID, [ref]$pageId, 1)

  if ([string]::IsNullOrWhiteSpace($pageId)) {
    throw "Failed to create page - pageId is empty"
  }

  Write-Host "Page created with ID: $pageId"

  # ページコンテンツ作成
  $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

  $pageContent = '<?xml version="1.0" encoding="UTF-8"?>'
  $pageContent += "<one:Page xmlns:one='$ns' ID='$pageId'>"
  $pageContent += "<one:Title><one:OEChildren><one:OE><one:T><![CDATA[$PageTitle]]></one:T></one:OE></one:OEChildren></one:Title>"
  $pageContent += "<one:Outline><one:OEChildren>"

  # 各行を追加
  $lines = $textContent -split "`n"
  $lineCount = 0
  foreach ($line in $lines) {
    if ($line.Trim() -ne '') {
      $pageContent += "<one:OE><one:T><![CDATA[$line]]></one:T></one:OE>"
      $lineCount++
    }
  }

  $pageContent += "</one:OEChildren></one:Outline>"
  $pageContent += "</one:Page>"

  Write-Host "Page content created: $lineCount lines"

  # OneNote に更新
  Write-Host "Updating OneNote page..."
  $one.UpdatePageContent($pageContent)
  Write-Host "Page updated successfully"

  Write-Host "SUCCESS"
  Write-Host "PageId=$pageId"
  Write-Host "Section=$SectionName"
  Write-Host "Lines=$lineCount"
}
catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  Write-Host "StackTrace: $($_.Exception.StackTrace)"
  exit 1
}
