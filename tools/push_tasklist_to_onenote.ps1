param(
  [string]$Date = (Get-Date -Format 'yyyy-MM-dd'),
  [string]$SectionName = '手書き日誌',
  [string]$PageTitle = '',
  [string]$VaultTaskListDir = "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳\12-TaskList",
  [string]$TaskHeading = '## タスク一覧'
)

$ErrorActionPreference = 'Stop'

$compactDate = $Date -replace '-', ''
if ([string]::IsNullOrWhiteSpace($PageTitle)) {
  $PageTitle = "${compactDate}_手書き文字起こし"
}

$taskFile = Join-Path $VaultTaskListDir "${Date}_Tasks.md"
if (-not (Test-Path -LiteralPath $taskFile -PathType Leaf)) {
  throw "Task file not found: $taskFile"
}

function Get-TaskKey([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  return ([regex]::Replace($Text.Trim().ToLowerInvariant(), '[\s,，、。．・]+', ''))
}

$lines = [System.IO.File]::ReadAllLines($taskFile, [System.Text.Encoding]::UTF8)
$incomingTasks = @()
foreach ($line in $lines) {
  if ($line -match '^\s*-\s*\[(?<c>[ xX])\]\s*(?<t>.+?)\s*$') {
    $completed = ($Matches['c'] -eq 'x' -or $Matches['c'] -eq 'X')
    $text = $Matches['t']
    if ($text -match '^(?<body>.+?)\s*\[(?<dev>[^\]]+)\]\s*$') {
      $text = $Matches['body']
    }
    $text = $text.Trim()
    if ($text) {
      $incomingTasks += [pscustomobject]@{ Text = $text; Completed = $completed }
    }
  }
}
if ($incomingTasks.Count -eq 0) {
  throw "No checkbox tasks parsed from: $taskFile"
}

$one = New-Object -ComObject OneNote.Application
$h = ''
$one.GetHierarchy('', 4, [ref]$h, 2)
[xml]$hierarchy = $h
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })
if ($sections.Count -ne 1) {
  throw "Section match count must be 1: $SectionName (found $($sections.Count))"
}
$section = $sections[0]

$sh = ''
$one.GetHierarchy($section.ID, 4, [ref]$sh, 2)
[xml]$sectionDoc = $sh
$pages = @($sectionDoc.SelectNodes('//*[local-name()="Page"]') | Where-Object { $_.name -eq $PageTitle })
if ($pages.Count -ne 1) {
  throw "Page match count must be 1: $SectionName / $PageTitle (found $($pages.Count))"
}
$pageId = $pages[0].ID

$pc = ''
$one.GetPageContent($pageId, [ref]$pc, 2)
[xml]$pageDoc = $pc

$tagDef = $pageDoc.SelectSingleNode('//*[local-name()="TagDef"][@name="To Do"]')
if ($null -eq $tagDef) {
  $tagDef = $pageDoc.CreateElement('one', 'TagDef', $pageDoc.DocumentElement.NamespaceURI)
  $tagDef.SetAttribute('index', '0')
  $tagDef.SetAttribute('type', '0')
  $tagDef.SetAttribute('symbol', '3')
  $tagDef.SetAttribute('fontColor', 'automatic')
  $tagDef.SetAttribute('highlightColor', 'none')
  $tagDef.SetAttribute('name', 'To Do')
  $firstChild = $pageDoc.DocumentElement.FirstChild
  [void]$pageDoc.DocumentElement.InsertBefore($tagDef, $firstChild)
}
$tagIndex = $tagDef.index

# Collect existing tagged tasks anywhere on the page (any earlier run's task
# section), remove them from their current position, and merge by normalized
# text so re-running never creates duplicates. Any task that only exists in
# OneNote (added by hand, not present in the source file) is preserved as-is.
$existingTasks = New-Object System.Collections.Specialized.OrderedDictionary
foreach ($oe in @($pageDoc.SelectNodes('//*[local-name()="OE"][*[local-name()="Tag"]]'))) {
  $tag = $oe.SelectSingleNode('./*[local-name()="Tag"]')
  $textNode = $oe.SelectSingleNode('./*[local-name()="T"]')
  $text = if ($null -ne $textNode) { $textNode.InnerText.Trim() } else { '' }
  $key = Get-TaskKey $text
  if ($key) {
    $existingTasks[$key] = [pscustomobject]@{
      Text = $text
      Completed = (([string]$tag.completed).ToLowerInvariant() -eq 'true')
    }
  }
  [void]$oe.ParentNode.RemoveChild($oe)
}
# Remove a prior heading OE for the task section too, so it isn't duplicated.
foreach ($oe in @($pageDoc.SelectNodes('//*[local-name()="OE"][*[local-name()="T"]]'))) {
  $textNode = $oe.SelectSingleNode('./*[local-name()="T"]')
  if ($null -ne $textNode -and $textNode.InnerText.Trim() -eq $TaskHeading.Trim()) {
    [void]$oe.ParentNode.RemoveChild($oe)
  }
}

$addedCount = 0
$updatedCount = 0
foreach ($task in $incomingTasks) {
  $key = Get-TaskKey $task.Text
  if (-not $key) { continue }
  if ($existingTasks.Contains($key)) {
    $existingTasks[$key] = [pscustomobject]@{ Text = $task.Text; Completed = $task.Completed }
    $updatedCount += 1
  } else {
    $existingTasks[$key] = [pscustomobject]@{ Text = $task.Text; Completed = $task.Completed }
    $addedCount += 1
  }
}

$oeChildren = $pageDoc.SelectNodes('//*[local-name()="Outline"]')[0].SelectSingleNode('./*[local-name()="OEChildren"]')

$headingOe = $pageDoc.CreateElement('one', 'OE', $pageDoc.DocumentElement.NamespaceURI)
$headingT = $pageDoc.CreateElement('one', 'T', $pageDoc.DocumentElement.NamespaceURI)
[void]$headingT.AppendChild($pageDoc.CreateCDataSection($TaskHeading))
[void]$headingOe.AppendChild($headingT)
[void]$oeChildren.AppendChild($headingOe)

foreach ($key in $existingTasks.Keys) {
  $task = $existingTasks[$key]
  $safeText = [string]$task.Text
  $safeText = $safeText.Replace('&', '＆').Replace('<', '＜').Replace('>', '＞').Replace('"', '”')
  $oe = $pageDoc.CreateElement('one', 'OE', $pageDoc.DocumentElement.NamespaceURI)
  $tag = $pageDoc.CreateElement('one', 'Tag', $pageDoc.DocumentElement.NamespaceURI)
  $tag.SetAttribute('index', $tagIndex)
  $tag.SetAttribute('completed', $(if ($task.Completed) { 'true' } else { 'false' }))
  [void]$oe.AppendChild($tag)
  $t = $pageDoc.CreateElement('one', 'T', $pageDoc.DocumentElement.NamespaceURI)
  [void]$t.AppendChild($pageDoc.CreateCDataSection($safeText))
  [void]$oe.AppendChild($t)
  [void]$oeChildren.AppendChild($oe)
}

$one.UpdatePageContent($pageDoc.OuterXml)

$doneCount = @($existingTasks.Values | Where-Object { $_.Completed }).Count
Write-Output "PageId=$pageId Added=$addedCount Updated=$updatedCount TotalTasks=$($existingTasks.Count) Done=$doneCount Open=$($existingTasks.Count - $doneCount)"
