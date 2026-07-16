param(
  [string]$NotebookName = 'FarEasternTribe',
  [string]$SectionName = '実験',
  [string]$DatePattern = '20260708',
  [string]$Out = '.\agent_workspace\実験ノートAgent\onenote_to_ppt\20260708_page_xml_debug.xml'
)

$ErrorActionPreference = 'Stop'
$one = New-Object -ComObject OneNote.Application
[xml]$hierarchy = ''
$one.GetHierarchy('', 4, [ref]$hierarchy)

$notebook = @($hierarchy.DocumentElement.SelectNodes('//*[local-name()="Notebook"]') | Where-Object {
  $_.name -eq $NotebookName
} | Select-Object -First 1)
if ($null -eq $notebook) { throw "Notebook not found: $NotebookName" }

$section = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
  $_.name -eq $SectionName
} | Select-Object -First 1)
if ($null -eq $section) { throw "Section not found: $SectionName" }

[xml]$sectionXml = ''
$one.GetHierarchy($section.ID, 4, [ref]$sectionXml)
$page = @($sectionXml.DocumentElement.SelectNodes('.//*[local-name()="Page"]') | Where-Object {
  $_.name -like "*$DatePattern*"
} | Select-Object -First 1)
if ($null -eq $page) { throw "Page not found: $DatePattern" }

$xmlText = ''
$one.GetPageContent($page.ID, [ref]$xmlText, 0)
[xml]$pageXml = $xmlText

$outPath = Resolve-Path -LiteralPath (Split-Path -Parent $Out) -ErrorAction SilentlyContinue
if ($null -eq $outPath) {
  New-Item -ItemType Directory -Path (Split-Path -Parent $Out) -Force | Out-Null
}
Set-Content -LiteralPath $Out -Value $xmlText -Encoding UTF8

Write-Host "Page: $($page.name)"
Write-Host "XML: $Out"
$pageXml.SelectNodes('//*') |
  Group-Object { $_.LocalName } |
  Sort-Object Count -Descending |
  Select-Object Count,Name |
  Format-Table -AutoSize

Write-Host "Media-like nodes:"
$pageXml.SelectNodes('//*[local-name()="Image" or local-name()="InsertedFile" or local-name()="InkDrawing" or local-name()="InkWord" or local-name()="MediaFile" or local-name()="Object" or local-name()="Data"]') |
  ForEach-Object {
    $attrs = @()
    foreach ($attr in $_.Attributes) {
      $attrs += "$($attr.Name)=$($attr.Value)"
    }
    Write-Host ("- " + $_.LocalName + " " + ($attrs -join ' '))
  }

