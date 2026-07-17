param(
  [string]$SectionName = '',
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$Out
)

$ErrorActionPreference = 'Stop'
$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText

$matches = @()
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]'))
foreach ($candidateSection in $sections) {
  if ($SectionName -and $candidateSection.name -ne $SectionName) { continue }

  $sectionText = ''
  $one.GetHierarchy($candidateSection.ID, 4, [ref]$sectionText, 2)
  [xml]$sectionXml = $sectionText
  foreach ($candidatePage in @($sectionXml.SelectNodes('//*[local-name()="Page"]') | Where-Object {
    $_.name -eq $PageTitle
  })) {
    $notebookName = ''
    $ancestor = $candidateSection.ParentNode
    while ($null -ne $ancestor) {
      if ($ancestor.LocalName -eq 'Notebook') {
        $notebookName = $ancestor.name
        break
      }
      $ancestor = $ancestor.ParentNode
    }
    $matches += [pscustomobject]@{
      NotebookName = $notebookName
      Section = $candidateSection
      Page = $candidatePage
    }
  }
}

if ($matches.Count -eq 0) {
  if ($SectionName) { throw "Page not found: $SectionName / $PageTitle" }
  throw "Page not found: $PageTitle"
}
if ($matches.Count -gt 1) {
  $locations = @($matches | ForEach-Object {
    "$($_.NotebookName) / $($_.Section.name) / $($_.Page.name)"
  }) -join '; '
  throw "Multiple exact-title pages found. Specify -SectionName. Matches: $locations"
}

$match = $matches[0]
$section = $match.Section
$page = $match.Page

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

# PublishFormat.pfPDF = 3.
$one.Publish($page.ID, $absoluteOut, 3, '')
if (-not (Test-Path -LiteralPath $absoluteOut)) { throw "PDF was not created: $absoluteOut" }
Write-Output "NotebookName=$($match.NotebookName)"
Write-Output "SectionName=$($section.name)"
Write-Output "Pdf=$absoluteOut"
