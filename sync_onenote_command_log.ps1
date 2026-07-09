param(
  [string]$NotebookName = 'OpenAI_Agent1',
  [string]$SectionName = '命令したLog',
  [string]$OutputDir = '.\agent_workspace\司令塔Agent\onenote_command_log'
)

$ErrorActionPreference = 'Stop'

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Workspace

try {
  $ResolvedOutputDir = Resolve-Path -LiteralPath $OutputDir -ErrorAction SilentlyContinue
  if ($null -eq $ResolvedOutputDir) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    $ResolvedOutputDir = Resolve-Path -LiteralPath $OutputDir
  }
  $OutputPath = $ResolvedOutputDir.Path

  $SnapshotDir = Join-Path $OutputPath 'pages'
  New-Item -ItemType Directory -Path $SnapshotDir -Force | Out-Null

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

  [xml]$sectionXml = ''
  $one.GetHierarchy($section.ID, 4, [ref]$sectionXml)

  $statePath = Join-Path $OutputPath 'state.json'
  $previousState = @{}
  if (Test-Path -LiteralPath $statePath) {
    try {
      $previousJson = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
      foreach ($item in @($previousJson.pages)) {
        if ($item.page_id) {
          $previousState[$item.page_id] = $item
        }
      }
    } catch {
      $previousState = @{}
    }
  }

  $pages = New-Object System.Collections.ArrayList
  $changedPages = New-Object System.Collections.ArrayList

  foreach ($page in @($sectionXml.DocumentElement.SelectNodes('.//*[local-name()="Page"]'))) {
    $pageXmlText = ''
    try {
      $one.GetPageContent($page.ID, [ref]$pageXmlText, 0)
    } catch {
      Write-Warning "Failed to read page: $($page.name) / $($_.Exception.Message)"
      continue
    }

    $plain = ''
    try {
      [xml]$pageXml = $pageXmlText
      $texts = @($pageXml.SelectNodes('//*[local-name()="T"]') | ForEach-Object { $_.InnerText })
      $plain = (($texts -join "`n") -replace "`r", '').Trim()
    } catch {
      $plain = ''
    }

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($plain)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
    $safeTitle = ($page.name -replace '[\\/:*?"<>|]', '_')
    if ($safeTitle.Length -gt 80) {
      $safeTitle = $safeTitle.Substring(0, 80)
    }
    $stamp = ''
    if ($page.lastModifiedTime) {
      try {
        $stamp = ([datetime]$page.lastModifiedTime).ToString('yyyyMMdd_HHmmss')
      } catch {
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
      }
    } else {
      $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    }

    $fileName = "${stamp}_${safeTitle}.md"
    $filePath = Join-Path $SnapshotDir $fileName
    $markdown = @(
      "# $($page.name)",
      "",
      "- Notebook: $NotebookName",
      "- Section: $SectionName",
      "- PageId: $($page.ID)",
      "- LastModified: $($page.lastModifiedTime)",
      "- Hash: $hash",
      "",
      "## Text",
      "",
      $plain
    ) -join "`n"
    Set-Content -LiteralPath $filePath -Value $markdown -Encoding UTF8

    $record = [ordered]@{
      page_id = [string]$page.ID
      title = [string]$page.name
      last_modified = [string]$page.lastModifiedTime
      hash = $hash
      markdown = $filePath
    }
    [void]$pages.Add($record)

    $previous = $previousState[[string]$page.ID]
    if ($null -eq $previous -or $previous.hash -ne $hash -or $previous.last_modified -ne [string]$page.lastModifiedTime) {
      [void]$changedPages.Add($record)
    }
  }

  $snapshot = [ordered]@{
    synced_at = (Get-Date).ToString('s')
    notebook = $NotebookName
    section = $SectionName
    page_count = $pages.Count
    changed_count = $changedPages.Count
    pages = @($pages)
  }
  $snapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8

  $changedPath = Join-Path $OutputPath 'changed_pages.json'
  @($changedPages) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $changedPath -Encoding UTF8

  $reportPath = Join-Path $OutputPath 'latest_sync_report.md'
  $reportLines = @(
    "# OneNote 命令したLog sync",
    "",
    "- Synced at: $($snapshot.synced_at)",
    "- Notebook: $NotebookName",
    "- Section: $SectionName",
    "- Page count: $($pages.Count)",
    "- Changed count: $($changedPages.Count)",
    "",
    "## Changed pages"
  )
  if ($changedPages.Count -eq 0) {
    $reportLines += "- No changes"
  } else {
    foreach ($item in @($changedPages)) {
      $reportLines += "- $($item.title) / $($item.last_modified) / $($item.markdown)"
    }
  }
  Set-Content -LiteralPath $reportPath -Value ($reportLines -join "`n") -Encoding UTF8

  Write-Host "Notebook: $NotebookName"
  Write-Host "Section: $SectionName"
  Write-Host "Pages: $($pages.Count)"
  Write-Host "Changed: $($changedPages.Count)"
  Write-Host "Report: $reportPath"
  Write-Host "State: $statePath"
} finally {
  Pop-Location
}




