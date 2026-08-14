param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile,
  [string]$ExistingPageId = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$resolvedTextFile = Resolve-Path -LiteralPath $TextFile
$textDirectory = Split-Path -Parent $resolvedTextFile
$lines = [System.IO.File]::ReadAllLines(
  $resolvedTextFile,
  [System.Text.Encoding]::UTF8
)

$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText
$sections = @($hierarchy.SelectNodes('//*[local-name()="Section"]') | Where-Object {
  $_.name -eq $SectionName
})
if ($sections.Count -ne 1) {
  throw "Section match count must be 1: $SectionName (found $($sections.Count))"
}
$section = $sections[0]

$sectionText = ''
$one.GetHierarchy($section.ID, 4, [ref]$sectionText, 2)
[xml]$sectionXml = $sectionText
$createdByThisRun = $false

function Remove-EmptyCreatedPage {
  param([Parameter(Mandatory=$true)][string]$CreatedPageId)
  try {
    $cleanupContent = ''
    $one.GetPageContent($CreatedPageId, [ref]$cleanupContent, 2)
    [xml]$cleanupDoc = $cleanupContent
    $outlineCount = @($cleanupDoc.SelectNodes('//*[local-name()="Outline"]')).Count
    $inkCount = @($cleanupDoc.SelectNodes('//*[local-name()="InkDrawing"]')).Count
    $imageCount = @($cleanupDoc.SelectNodes('//*[local-name()="Image"]')).Count
    if ($outlineCount -eq 0 -and $inkCount -eq 0 -and $imageCount -eq 0) {
      $one.DeleteHierarchy($CreatedPageId, 0)
      Write-Warning "Removed empty OneNote page created by failed write: $CreatedPageId"
    } else {
      Write-Warning "Kept partially populated OneNote page for manual review: $CreatedPageId"
    }
  } catch {
    Write-Warning "Could not verify or remove failed OneNote page $CreatedPageId`: $($_.Exception.Message)"
  }
}

try {
  $pageId = $ExistingPageId
  if (-not [string]::IsNullOrWhiteSpace($pageId)) {
    $sectionPageIds = @(
      $sectionXml.SelectNodes('//*[local-name()="Page"]') |
        ForEach-Object { $_.GetAttribute('ID') }
    )
    if ($sectionPageIds -notcontains $pageId) {
      throw "Existing page does not belong to section '$SectionName': $pageId"
    }
  }
  if ([string]::IsNullOrWhiteSpace($pageId)) {
  $baseTitle = $PageTitle
  $suffix = 2
  while (@($sectionXml.SelectNodes('//*[local-name()="Page"]') | Where-Object {
    $_.name -eq $PageTitle
  }).Count -gt 0) {
    $PageTitle = "${baseTitle}_$suffix"
    $suffix += 1
  }
  $one.CreateNewPage($section.ID, [ref]$pageId, 1)
    $createdByThisRun = $true
  }
  $namespace = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
  $existingContent = ''
  $one.GetPageContent($pageId, [ref]$existingContent, 2)
  [xml]$doc = $existingContent
  $page = $doc.DocumentElement
  $existingOutlineCount = @($page.SelectNodes('./*[local-name()="Outline"]')).Count
  $existingInkCount = @($page.SelectNodes('.//*[local-name()="InkDrawing"]')).Count
  $existingImageCount = @($page.SelectNodes('.//*[local-name()="Image"]')).Count
  $existingTitleText = [string]$page.SelectSingleNode(
    './*[local-name()="Title"]/*[local-name()="OE"]/*[local-name()="T"]'
  ).InnerText
  if (
    $existingOutlineCount -gt 0 -or
    $existingInkCount -gt 0 -or
    $existingImageCount -gt 0 -or
    (-not [string]::IsNullOrWhiteSpace($existingTitleText))
  ) {
    throw "Refusing to replace a non-empty existing page: $pageId"
  }
  $titleText = $page.SelectSingleNode(
    './*[local-name()="Title"]/*[local-name()="OE"]/*[local-name()="T"]'
  )
  if ($null -eq $titleText) {
    throw "OneNote page title node was not found: $pageId"
  }
  $titleText.RemoveAll()
  [void]$titleText.AppendChild($doc.CreateCDataSection($PageTitle))

$outline = $doc.CreateElement('one', 'Outline', $namespace)
$position = $doc.CreateElement('one', 'Position', $namespace)
$position.SetAttribute('x', '36')
$position.SetAttribute('y', '86')
$position.SetAttribute('z', '0')
[void]$outline.AppendChild($position)
$children = $doc.CreateElement('one', 'OEChildren', $namespace)

  $imageCount = 0
  $figureCount = 0
  $equationCount = 0
  $textLineCount = 0
  $expectedTextLines = New-Object System.Collections.Generic.List[string]
  $expectedVisualKinds = New-Object System.Collections.Generic.List[string]
  foreach ($line in $lines) {
  if ($line -match '^\[\[(?<kind>IMAGE|FIGURE|EQUATION):(?<path>.+?)(?:\|(?<alt>.*?))?\]\]$') {
    $visualKind = $Matches.kind.ToUpperInvariant()
    $imagePath = $Matches.path
    $altText = if ($Matches.ContainsKey('alt')) { $Matches.alt } else { '' }
    if (-not [System.IO.Path]::IsPathRooted($imagePath)) {
      $imagePath = Join-Path $textDirectory $imagePath
    }
    $resolvedImage = Resolve-Path -LiteralPath $imagePath
    $extension = [System.IO.Path]::GetExtension($resolvedImage).ToLowerInvariant()
    $format = switch ($extension) {
      '.jpg' { 'jpg' }
      '.jpeg' { 'jpg' }
      '.png' { 'png' }
      '.emf' { 'emf' }
      default { throw "Unsupported OneNote image format: $extension ($resolvedImage)" }
    }
    $bytes = [System.IO.File]::ReadAllBytes($resolvedImage)
    $imageInfo = [System.Drawing.Image]::FromFile($resolvedImage)
    try {
      $maxWidth = 650.0
      $width = [double]$imageInfo.Width
      $height = [double]$imageInfo.Height
      if ($width -gt $maxWidth) {
        $height = $height * ($maxWidth / $width)
        $width = $maxWidth
      }
    }
    finally {
      $imageInfo.Dispose()
    }

    $oe = $doc.CreateElement('one', 'OE', $namespace)
    $image = $doc.CreateElement('one', 'Image', $namespace)
    $image.SetAttribute('format', $format)
    $typedAlt = if ([string]::IsNullOrWhiteSpace($altText)) {
      "handwriting-$($visualKind.ToLowerInvariant())"
    } else {
      "handwriting-$($visualKind.ToLowerInvariant()): $altText"
    }
    $image.SetAttribute('alt', $typedAlt)
    $size = $doc.CreateElement('one', 'Size', $namespace)
    $size.SetAttribute('width', $width.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture))
    $size.SetAttribute('height', $height.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture))
    [void]$image.AppendChild($size)
    $data = $doc.CreateElement('one', 'Data', $namespace)
    [void]$data.AppendChild($doc.CreateTextNode([Convert]::ToBase64String($bytes)))
    [void]$image.AppendChild($data)
    [void]$oe.AppendChild($image)
    [void]$children.AppendChild($oe)
    $imageCount += 1
    [void]$expectedVisualKinds.Add($visualKind)
    if ($visualKind -eq 'FIGURE') { $figureCount += 1 }
    if ($visualKind -eq 'EQUATION') { $equationCount += 1 }
    continue
  }

    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $oe = $doc.CreateElement('one', 'OE', $namespace)
    $text = $doc.CreateElement('one', 'T', $namespace)
    [void]$text.AppendChild($doc.CreateCDataSection($line))
    [void]$oe.AppendChild($text)
    [void]$children.AppendChild($oe)
    [void]$expectedTextLines.Add($line)
    $textLineCount += 1
  }

  if ($textLineCount -eq 0 -and $imageCount -eq 0) {
    throw 'Refusing to create an empty OneNote page.'
  }

[void]$outline.AppendChild($children)
[void]$page.AppendChild($outline)
$updated = $false
for ($attempt = 1; $attempt -le 5; $attempt++) {
  try {
    $one.UpdatePageContent($doc.OuterXml)
    $updated = $true
    break
  } catch {
    if ($attempt -eq 5) {
      $hresult = '{0:X8}' -f ($_.Exception.HResult -band 0xFFFFFFFFL)
      if ($hresult -eq '80042030') {
        throw 'OneNote is blocked by a modal dialog (HRESULT 0x80042030). Dismiss the visible OneNote dialog and retry; no content was written.'
      }
      throw
    }
    Start-Sleep -Milliseconds (500 * $attempt)
  }
}
if (-not $updated) { throw 'OneNote page update did not complete.' }

$readBack = ''
$one.GetPageContent($pageId, [ref]$readBack, 2)
[xml]$readDoc = $readBack
$readTitle = [string]$readDoc.Page.Title.OE.T.'#cdata-section'
$readTextNodes = @($readDoc.SelectNodes('//*[local-name()="Outline"]//*[local-name()="T"]'))
$readImageNodes = @($readDoc.SelectNodes('//*[local-name()="Outline"]//*[local-name()="Image"]'))
$readLines = @($readTextNodes | ForEach-Object { [string]$_.InnerText })
$readVisualKinds = @(
  $readImageNodes | ForEach-Object {
    $alt = [string]$_.GetAttribute('alt')
    if ($alt -match '^handwriting-(image|figure|equation)(?::|$)') {
      $Matches[1].ToUpperInvariant()
    } else {
      'IMAGE'
    }
  }
)

Write-Output "SectionName=$SectionName"
Write-Output "PageTitle=$readTitle"
Write-Output "PageId=$pageId"
Write-Output "TextLines=$($readLines.Count)"
Write-Output "ExpectedTextLines=$textLineCount"
Write-Output "Images=$($readImageNodes.Count)"
Write-Output "ExpectedImages=$imageCount"
Write-Output "Figures=$figureCount"
Write-Output "Equations=$equationCount"
Write-Output "VisualOrder=$($readVisualKinds -join ',')"
Write-Output "FirstLine=$($readLines | Select-Object -First 1)"
Write-Output "LastLine=$($readLines | Select-Object -Last 1)"

  if ($readTitle -cne $PageTitle) { throw "OneNote title read-back mismatch." }
  if ($readLines.Count -ne $textLineCount) { throw "OneNote text-line count mismatch." }
  for ($index = 0; $index -lt $expectedTextLines.Count; $index++) {
    if ($readLines[$index] -cne $expectedTextLines[$index]) {
      throw "OneNote text read-back mismatch at line $($index + 1)."
    }
  }
  if ($readImageNodes.Count -ne $imageCount) { throw "OneNote image count mismatch." }
  if ($readVisualKinds.Count -ne $expectedVisualKinds.Count) {
    throw "OneNote visual-kind count mismatch."
  }
  for ($index = 0; $index -lt $expectedVisualKinds.Count; $index++) {
    if ($readVisualKinds[$index] -cne $expectedVisualKinds[$index]) {
      throw "OneNote visual-kind read-back mismatch at visual $($index + 1)."
    }
  }

  $automationOutputRegistered = $false
  $transcriptionSuffix = [System.Text.Encoding]::UTF8.GetString(
    [System.Convert]::FromBase64String('X+aJi+abuOOBjeaWh+Wtl+i1t+OBk+OBlw==')
  )
  $isTranscriptionOutput = $readTitle -match '^20\d{6}_' -and
    $readTitle.Substring(8).StartsWith($transcriptionSuffix, [System.StringComparison]::Ordinal)
  if ($createdByThisRun -and $isTranscriptionOutput) {
    $workspace = Split-Path -Parent $PSScriptRoot
    $registryDir = Join-Path $workspace '.agent_runtime\transcription_outputs'
    New-Item -ItemType Directory -Path $registryDir -Force | Out-Null
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
      $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($pageId))
    }
    finally {
      $sha.Dispose()
    }
    $registryKey = ([System.BitConverter]::ToString($hashBytes) -replace '-', '').Substring(0, 16).ToLowerInvariant()
    $lastModifiedTime = [string]$readDoc.DocumentElement.GetAttribute('lastModifiedTime')
    if ([string]::IsNullOrWhiteSpace($lastModifiedTime)) {
      $lastModifiedTime = [datetime]::UtcNow.ToString('o')
    }
    $contentSha = [System.Security.Cryptography.SHA256]::Create()
    try {
      $contentHashBytes = $contentSha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($readBack))
    }
    finally {
      $contentSha.Dispose()
    }
    $contentHash = ([System.BitConverter]::ToString($contentHashBytes) -replace '-', '').ToLowerInvariant()
    $registryRecord = [ordered]@{
      PageId = $pageId
      PageTitle = $readTitle
      LastModifiedTime = $lastModifiedTime
      ContentHash = $contentHash
      RegisteredAt = [datetime]::UtcNow.ToString('o')
    }
    $registryPath = Join-Path $registryDir "$registryKey.json"
    $registryRecord | ConvertTo-Json | Set-Content -LiteralPath $registryPath -Encoding UTF8
    $automationOutputRegistered = $true
  }
  Write-Output "AutomationOutputRegistered=$($automationOutputRegistered.ToString().ToLowerInvariant())"
} catch {
  if ($createdByThisRun -and -not [string]::IsNullOrWhiteSpace($pageId)) {
    Remove-EmptyCreatedPage -CreatedPageId $pageId
  }
  throw
}
