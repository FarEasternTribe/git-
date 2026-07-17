param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Stop'
$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText
$section = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object {
  $_.name -eq $SectionName
} | Select-Object -First 1)
if ($null -eq $section) { throw "Section not found: $SectionName" }

$sectionText = ''
$one.GetHierarchy($section.ID, 4, [ref]$sectionText, 2)
[xml]$sectionXml = $sectionText
$baseTitle = $PageTitle
$suffix = 2
while (@($sectionXml.SelectNodes('//*[local-name()="Page"]') | Where-Object { $_.name -eq $PageTitle }).Count -gt 0) {
  $PageTitle = "${baseTitle}_$suffix"
  $suffix += 1
}

$text = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $TextFile), [System.Text.Encoding]::UTF8)
$pageId = ''
$one.CreateNewPage($section.ID, [ref]$pageId, 0)
$namespace = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$doc = New-Object System.Xml.XmlDocument
$page = $doc.CreateElement('one', 'Page', $namespace)
$page.SetAttribute('ID', $pageId)
[void]$doc.AppendChild($page)
$title = $doc.CreateElement('one', 'Title', $namespace)
$titleOE = $doc.CreateElement('one', 'OE', $namespace)
$titleText = $doc.CreateElement('one', 'T', $namespace)
[void]$titleText.AppendChild($doc.CreateCDataSection($PageTitle))
[void]$titleOE.AppendChild($titleText)
[void]$title.AppendChild($titleOE)
[void]$page.AppendChild($title)
$outline = $doc.CreateElement('one', 'Outline', $namespace)
$position = $doc.CreateElement('one', 'Position', $namespace)
$position.SetAttribute('x', '36')
$position.SetAttribute('y', '86')
$position.SetAttribute('z', '0')
[void]$outline.AppendChild($position)
$children = $doc.CreateElement('one', 'OEChildren', $namespace)
foreach ($line in ($text -replace "`r`n", "`n" -split "`n")) {
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  $oe = $doc.CreateElement('one', 'OE', $namespace)
  $t = $doc.CreateElement('one', 'T', $namespace)
  [void]$t.AppendChild($doc.CreateCDataSection($line))
  [void]$oe.AppendChild($t)
  [void]$children.AppendChild($oe)
}
[void]$outline.AppendChild($children)
[void]$page.AppendChild($outline)
$one.UpdatePageContent($doc.OuterXml)

$readBack = ''
$one.GetPageContent($pageId, [ref]$readBack, 2)
[xml]$readDoc = $readBack
$lineCount = @($readDoc.SelectNodes('//*[local-name()="Outline"]//*[local-name()="T"]')).Count
Write-Output "PageTitle=$PageTitle"
Write-Output "PageId=$pageId"
Write-Output "TextLines=$lineCount"
