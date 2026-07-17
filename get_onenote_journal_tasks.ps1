param(
  [string]$NotebookName = 'FarEasternTribe',
  [string]$SectionName = '',
  [string]$Date = ([DateTime]::Now.ToString('yyyy-MM-dd')),
  [string]$PageTitle = '',
  [switch]$Json
)

$ErrorActionPreference = 'Stop'
$journalLabel = -join ([char]0x65E5, [char]0x8A8C)
if ([string]::IsNullOrWhiteSpace($SectionName)) { $SectionName = $journalLabel }

function Get-PlainText([string]$Value) {
  if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
  $decoded = [System.Net.WebUtility]::HtmlDecode($Value)
  return ([regex]::Replace($decoded, '<[^>]+>', '')).Trim()
}

$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText

$notebook = @($hierarchy.SelectNodes('//*[local-name()="Notebook"]') | Where-Object {
  $_.name -eq $NotebookName
} | Select-Object -First 1)
if ($null -eq $notebook) { throw "Notebook not found: $NotebookName" }

$section = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
  $_.name -eq $SectionName
} | Select-Object -First 1)
if ($null -eq $section) { throw "Section not found: $SectionName" }

$sectionText = ''
$one.GetHierarchy($section.ID, 4, [ref]$sectionText, 2)
[xml]$sectionXml = $sectionText
$targetTitle = if ([string]::IsNullOrWhiteSpace($PageTitle)) { "${Date}_$journalLabel" } else { $PageTitle }
$page = @($sectionXml.SelectNodes('//*[local-name()="Page"]') | Where-Object {
  $_.name -eq $targetTitle
} | Select-Object -First 1)
if ($null -eq $page) {
  $page = @($sectionXml.SelectNodes('//*[local-name()="Page"]') | Where-Object {
    $_.name -like "*$Date*"
  } | Select-Object -First 1)
}
if ($null -eq $page) { throw "Journal page not found: $targetTitle" }

$pageText = ''
$one.GetPageContent($page.ID, [ref]$pageText, 2)
[xml]$pageXml = $pageText

$tagDefinitions = @{}
foreach ($definition in $pageXml.SelectNodes('//*[local-name()="TagDef"]')) {
  $tagDefinitions[[string]$definition.index] = [string]$definition.name
}

$tasks = @()
foreach ($oe in $pageXml.SelectNodes('//*[local-name()="Outline"]//*[local-name()="OE"]')) {
  $tag = $oe.SelectSingleNode('./*[local-name()="Tag"]')
  if ($null -eq $tag) { continue }
  $tagName = $tagDefinitions[[string]$tag.index]
  if ($tagName -notmatch '^(To Do|Todo)') { continue }
  $textNode = $oe.SelectSingleNode('./*[local-name()="T"]')
  $taskText = if ($null -eq $textNode) { '' } else { Get-PlainText $textNode.InnerText }
  if ([string]::IsNullOrWhiteSpace($taskText)) { continue }
  $isCompleted = ([string]$tag.completed).ToLowerInvariant() -eq 'true'
  $tasks += [pscustomobject]@{
    text = $taskText
    completed = $isCompleted
  }
}

$remaining = @($tasks | Where-Object { -not $_.completed })
$completed = @($tasks | Where-Object { $_.completed })
$result = [pscustomobject]@{
  page = [string]$page.name
  date = $Date
  total = $tasks.Count
  remaining_count = $remaining.Count
  completed_count = $completed.Count
  remaining = @($remaining.text)
  completed = @($completed.text)
}

if ($Json) {
  $result | ConvertTo-Json -Depth 4
  exit 0
}

Write-Output "# OneNote journal task status"
Write-Output ""
Write-Output "- Page: $($result.page)"
Write-Output "- Total: $($result.total)"
Write-Output "- Remaining: $($result.remaining_count)"
Write-Output "- Completed: $($result.completed_count)"
Write-Output ""
Write-Output "## Remaining tasks"
if ($remaining.Count -eq 0) { Write-Output '- None' }
foreach ($task in $remaining) { Write-Output "- [ ] $($task.text)" }
Write-Output ""
Write-Output "## Completed tasks"
if ($completed.Count -eq 0) { Write-Output '- None' }
foreach ($task in $completed) { Write-Output "- [x] $($task.text)" }
