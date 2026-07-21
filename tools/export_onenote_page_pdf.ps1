param(
  [string]$SectionName = '',
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$Out,
  [switch]$LatestForDate
)

$ErrorActionPreference = 'Stop'
$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText

$workspace = Split-Path -Parent $PSScriptRoot
$outputRegistryDir = Join-Path $workspace '.agent_runtime\transcription_outputs'
$registeredOutputs = @{}
if (Test-Path -LiteralPath $outputRegistryDir -PathType Container) {
  foreach ($registryFile in @(Get-ChildItem -LiteralPath $outputRegistryDir -Filter '*.json' -File)) {
    try {
      $record = Get-Content -LiteralPath $registryFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
      if (-not [string]::IsNullOrWhiteSpace([string]$record.PageId)) {
        $registeredOutputs[[string]$record.PageId] = $record
      }
    }
    catch {
      Write-Warning "Ignored invalid transcription-output registry: $($registryFile.FullName)"
    }
  }
}

if ($LatestForDate -and $PageTitle -notmatch '^20\d{6}$') {
  throw '-LatestForDate requires PageTitle in YYYYMMDD form.'
}
$requestedTitle = $PageTitle
$latestDateMode = [bool]$LatestForDate
$transcriptionSuffix = [System.Text.Encoding]::UTF8.GetString(
  [System.Convert]::FromBase64String('X+aJi+abuOOBjeaWh+Wtl+i1t+OBk+OBlw==')
)
$transcriptionPrefix = "${PageTitle}${transcriptionSuffix}"

function Test-UntouchedAutomationOutput {
  param([Parameter(Mandatory=$true)]$CandidatePage)
  $pageId = [string]$CandidatePage.ID
  if (-not $script:registeredOutputs.ContainsKey($pageId)) { return $false }
  $record = $script:registeredOutputs[$pageId]
  $registeredModified = [string]$record.LastModifiedTime
  $currentModified = [string]$CandidatePage.lastModifiedTime
  if ([string]::IsNullOrWhiteSpace($registeredModified) -or
      [string]::IsNullOrWhiteSpace($currentModified)) {
    return $false
  }
  try {
    $registeredTime = [datetime]::Parse(
      $registeredModified,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::RoundtripKind
    )
    $currentTime = [datetime]::Parse(
      $currentModified,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::RoundtripKind
    )
    if ($currentTime -gt $registeredTime) { return $false }
    $registeredHash = [string]$record.ContentHash
    if ([string]::IsNullOrWhiteSpace($registeredHash)) { return $true }
    try {
      $pageContent = ''
      $script:one.GetPageContent($pageId, [ref]$pageContent, 2)
      $sha = [System.Security.Cryptography.SHA256]::Create()
      try {
        $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($pageContent))
      }
      finally {
        $sha.Dispose()
      }
      $currentHash = ([System.BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
      return $currentHash -ceq $registeredHash
    }
    catch {
      return $false
    }
  }
  catch {
    return $false
  }
}

function Select-TargetMatch {
  param([Parameter(Mandatory=$true)][object[]]$Candidates)
  if (-not $script:LatestForDate) {
    if ($Candidates.Count -gt 1) {
      $locations = @($Candidates | ForEach-Object {
        "$($_.NotebookName) / $($_.Section.name) / $($_.Page.name)"
      }) -join '; '
      throw "Multiple exact-title pages found. Specify -SectionName. Matches: $locations"
    }
    return $Candidates[0]
  }

  $eligible = @($Candidates | Where-Object { -not (Test-UntouchedAutomationOutput $_.Page) })
  if ($eligible.Count -eq 0) {
    throw "No eligible latest page found for date: $($script:PageTitle)"
  }
  return @($eligible | Sort-Object @{
    Expression = {
      try { [datetime]::Parse([string]$_.Page.lastModifiedTime) }
      catch { [datetime]::MinValue }
    }
    Descending = $true
  }, @{
    Expression = { [string]$_.Page.name }
    Descending = $true
  })[0]
}

$matches = @()
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]'))
foreach ($candidateSection in $sections) {
  if ($SectionName -and $candidateSection.name -ne $SectionName) { continue }

  $sectionText = ''
  $one.GetHierarchy($candidateSection.ID, 4, [ref]$sectionText, 2)
  [xml]$sectionXml = $sectionText
  foreach ($candidatePage in @($sectionXml.SelectNodes('//*[local-name()="Page"]'))) {
    $candidateTitle = [string]$candidatePage.name
    $titleMatches = if ($latestDateMode) {
      ($candidateTitle -ceq $requestedTitle) -or
        $candidateTitle.StartsWith($transcriptionPrefix, [System.StringComparison]::Ordinal)
    }
    else {
      $candidateTitle -ceq $requestedTitle
    }
    if (-not $titleMatches) { continue }
    $notebookName = ''
    $notebookId = ''
    $ancestor = $candidateSection.ParentNode
    while ($null -ne $ancestor) {
      if ($ancestor.LocalName -eq 'Notebook') {
        $notebookName = $ancestor.name
        $notebookId = $ancestor.ID
        break
      }
      $ancestor = $ancestor.ParentNode
    }
    $matches += [pscustomobject]@{
      NotebookName = $notebookName
      NotebookId = $notebookId
      Section = $candidateSection
      Page = $candidatePage
    }
  }
}

if ($matches.Count -eq 0) {
  if ($SectionName) { throw "Page not found: $SectionName / $PageTitle" }
  throw "Page not found: $PageTitle"
}
$match = Select-TargetMatch -Candidates $matches
$section = $match.Section
$page = $match.Page
$syncRequested = $false

# OneNote may expose an older local snapshot until its notebook hierarchy is
# explicitly synchronized. Refresh the target notebook before publishing and
# wait for the page metadata to become stable across two consecutive reads.
if (-not [string]::IsNullOrWhiteSpace($match.NotebookId)) {
  try {
    $one.SyncHierarchy($match.NotebookId)
    $syncRequested = $true
  }
  catch {
    Write-Warning "OneNote notebook sync request failed: $($_.Exception.Message)"
  }
}

$stableReads = 0
$previousModified = ''
for ($attempt = 0; $attempt -lt 5; $attempt++) {
  Start-Sleep -Seconds 2
  $sectionText = ''
  $one.GetHierarchy($section.ID, 4, [ref]$sectionText, 2)
  [xml]$sectionXml = $sectionText
  $refreshedMatches = @()
  foreach ($refreshedPage in @($sectionXml.SelectNodes('//*[local-name()="Page"]'))) {
    $refreshedTitle = [string]$refreshedPage.name
    $titleMatches = if ($latestDateMode) {
      ($refreshedTitle -ceq $requestedTitle) -or
        $refreshedTitle.StartsWith($transcriptionPrefix, [System.StringComparison]::Ordinal)
    }
    else {
      $refreshedTitle -ceq $requestedTitle
    }
    if (-not $titleMatches) { continue }
    $refreshedMatches += [pscustomobject]@{
      NotebookName = $match.NotebookName
      NotebookId = $match.NotebookId
      Section = $section
      Page = $refreshedPage
    }
  }
  if ($refreshedMatches.Count -eq 0) {
    throw "Page changed during synchronization: $PageTitle"
  }
  $match = Select-TargetMatch -Candidates $refreshedMatches
  $page = $match.Page
  $currentModified = [string]$page.lastModifiedTime
  if ($currentModified -eq $previousModified -and -not [string]::IsNullOrWhiteSpace($currentModified)) {
    $stableReads += 1
    if ($stableReads -ge 1) { break }
  }
  else {
    $stableReads = 0
  }
  $previousModified = $currentModified
}

if ([System.IO.Path]::IsPathRooted($Out)) {
  $absoluteOut = [System.IO.Path]::GetFullPath($Out)
}
else {
  $absoluteOut = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Out))
}
$parent = Split-Path -Parent $absoluteOut
if (-not (Test-Path -LiteralPath $parent)) {
  New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
if (Test-Path -LiteralPath $absoluteOut -PathType Leaf) {
  Remove-Item -LiteralPath $absoluteOut -Force
}

# PublishFormat.pfPDF = 3.
$one.Publish($page.ID, $absoluteOut, 3, '')
if (-not (Test-Path -LiteralPath $absoluteOut)) { throw "PDF was not created: $absoluteOut" }
Write-Output "NotebookName=$($match.NotebookName)"
Write-Output "SectionName=$($section.name)"
Write-Output "SelectionMode=$(if ($LatestForDate) { 'LatestForDate' } else { 'ExactTitle' })"
Write-Output "LatestDateMode=$latestDateMode"
Write-Output "RequestedPageTitle=$PageTitle"
Write-Output "ResolvedPageTitle=$($page.name)"
Write-Output "ResolvedPageId=$($page.ID)"
Write-Output "CandidateCount=$($matches.Count)"
Write-Output "MatchedPageTitles=$(@($matches | ForEach-Object { [string]$_.Page.name }) -join '|')"
Write-Output "PageLastModified=$($page.lastModifiedTime)"
Write-Output "SyncRequested=$syncRequested"
Write-Output "Pdf=$absoluteOut"
