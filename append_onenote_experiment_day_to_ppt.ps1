param(
  [string]$Date = (Get-Date).ToString('yyyy-MM-dd'),
  [string]$NotebookName = '2026実験',
  [string]$SectionName = '実験',
  [string]$Pptx = '.\Experiment.pptx',
  [string]$Device = '',
  [string]$OutboxDir = '.\agent_workspace\実験ノートAgent\onenote_to_ppt',
  [switch]$Force,
  [switch]$ReplaceDateSlides,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Workspace

function Normalize-Text([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  $value = [System.Net.WebUtility]::HtmlDecode($Text)
  $value = $value -replace '(?i)<br\s*/?>', "`n"
  $value = $value -replace '(?is)</p\s*>', "`n"
  $value = $value -replace '(?is)</div\s*>', "`n"
  $value = $value -replace '(?is)<[^>]+>', ''
  $value = $value -replace "`r", ''
  $value = $value -replace "[`t ]+`n", "`n"
  $value = $value -replace "`n[`t ]+", "`n"
  $value = $value -replace "`n{3,}", "`n`n"
  return $value.Trim()
}

function Get-DeviceLabel {
  if (-not [string]::IsNullOrWhiteSpace($Device)) { return $Device }
  $computerName = [string]$env:COMPUTERNAME
  if ($computerName.ToUpperInvariant().Contains('LENOVO')) { return 'Lenovo' }
  if ($computerName.ToUpperInvariant().Contains('DESKTOP')) { return 'Desktop' }
  if (-not [string]::IsNullOrWhiteSpace($computerName)) { return $computerName }
  return 'UnknownPC'
}

function Split-TextChunks([string]$Text, [int]$MaxChars = 1100) {
  $clean = Normalize-Text $Text
  if ([string]::IsNullOrWhiteSpace($clean)) { return @('') }
  $chunks = New-Object System.Collections.ArrayList
  $current = ''
  foreach ($line in ($clean -split "`n")) {
    $candidate = if ($current) { "$current`n$line" } else { $line }
    if ($candidate.Length -gt $MaxChars -and $current) {
      [void]$chunks.Add($current)
      $current = $line
    } else {
      $current = $candidate
    }
  }
  if ($current) { [void]$chunks.Add($current) }
  return @($chunks)
}

function Get-DatePatterns([datetime]$DateValue) {
  return @(
    $DateValue.ToString('yyyy-MM-dd'),
    $DateValue.ToString('yyyy/MM/dd'),
    $DateValue.ToString('yyyy.M.d'),
    $DateValue.ToString('yyyyMMdd'),
    ($DateValue.ToString('yyyy年M月d日')),
    ($DateValue.ToString('M月d日'))
  )
}

function Get-SectionPath($Section) {
  $parts = New-Object System.Collections.ArrayList
  $node = $Section
  while ($null -ne $node -and $node.LocalName -ne 'Notebook') {
    if (($node.LocalName -eq 'Section' -or $node.LocalName -eq 'SectionGroup') -and
        -not [string]::IsNullOrWhiteSpace($node.name)) {
      [void]$parts.Insert(0, [string]$node.name)
    }
    $node = $node.ParentNode
  }
  return ($parts -join ' / ')
}

function Add-TextSlide($Presentation, [string]$Title, [string]$Meta, [string]$Body) {
  $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
  $width = $Presentation.PageSetup.SlideWidth
  $height = $Presentation.PageSetup.SlideHeight
  $margin = 32

  $titleBox = $slide.Shapes.AddTextbox(1, $margin, 20, $width - ($margin * 2), 44)
  $titleBox.TextFrame.TextRange.Text = $Title
  $titleBox.TextFrame.TextRange.Font.Size = 24
  $titleBox.TextFrame.TextRange.Font.Bold = $true

  $metaBox = $slide.Shapes.AddTextbox(1, $margin, 64, $width - ($margin * 2), 28)
  $metaBox.TextFrame.TextRange.Text = $Meta
  $metaBox.TextFrame.TextRange.Font.Size = 10
  $metaBox.TextFrame.TextRange.Font.Color.RGB = 0x666666

  $bodyBox = $slide.Shapes.AddTextbox(1, $margin, 102, $width - ($margin * 2), $height - 130)
  $bodyBox.TextFrame.TextRange.Text = $Body
  $bodyBox.TextFrame.TextRange.Font.Size = 12
  $bodyBox.TextFrame.MarginLeft = 8
  $bodyBox.TextFrame.MarginRight = 8
  $bodyBox.TextFrame.MarginTop = 6
  $bodyBox.TextFrame.MarginBottom = 6

  return $slide
}

function Add-ImageSlide($Presentation, [string]$Title, [string]$Meta, [string]$ImagePath) {
  $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
  $width = $Presentation.PageSetup.SlideWidth
  $height = $Presentation.PageSetup.SlideHeight
  $margin = 32

  $titleBox = $slide.Shapes.AddTextbox(1, $margin, 20, $width - ($margin * 2), 44)
  $titleBox.TextFrame.TextRange.Text = $Title
  $titleBox.TextFrame.TextRange.Font.Size = 24
  $titleBox.TextFrame.TextRange.Font.Bold = $true

  $metaBox = $slide.Shapes.AddTextbox(1, $margin, 64, $width - ($margin * 2), 28)
  $metaBox.TextFrame.TextRange.Text = $Meta
  $metaBox.TextFrame.TextRange.Font.Size = 10
  $metaBox.TextFrame.TextRange.Font.Color.RGB = 0x666666

  $shape = $slide.Shapes.AddPicture($ImagePath, $false, $true, $margin, 104, -1, -1)
  $maxW = $width - ($margin * 2)
  $maxH = $height - 132
  $scale = [Math]::Min($maxW / $shape.Width, $maxH / $shape.Height)
  if ($scale -lt 1) {
    $shape.Width = $shape.Width * $scale
    $shape.Height = $shape.Height * $scale
  }
  $shape.Left = ($width - $shape.Width) / 2
  $shape.Top = 104 + (($maxH - $shape.Height) / 2)
  return $slide
}

function Add-TableSlide($Presentation, [string]$Title, [string]$Meta, [object[]]$Rows) {
  $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
  $width = $Presentation.PageSetup.SlideWidth
  $height = $Presentation.PageSetup.SlideHeight
  $margin = 32

  $titleBox = $slide.Shapes.AddTextbox(1, $margin, 20, $width - ($margin * 2), 44)
  $titleBox.TextFrame.TextRange.Text = $Title
  $titleBox.TextFrame.TextRange.Font.Size = 24
  $titleBox.TextFrame.TextRange.Font.Bold = $true

  $metaBox = $slide.Shapes.AddTextbox(1, $margin, 64, $width - ($margin * 2), 28)
  $metaBox.TextFrame.TextRange.Text = $Meta
  $metaBox.TextFrame.TextRange.Font.Size = 10
  $metaBox.TextFrame.TextRange.Font.Color.RGB = 0x666666

  $rowCount = [Math]::Max(1, @($Rows).Count)
  $colCount = 1
  foreach ($row in @($Rows)) {
    $colCount = [Math]::Max($colCount, @($row).Count)
  }
  $tableShape = $slide.Shapes.AddTable($rowCount, $colCount, $margin, 104, $width - ($margin * 2), $height - 132)
  $table = $tableShape.Table
  for ($r = 1; $r -le $rowCount; $r++) {
    $row = @($Rows[$r - 1])
    for ($c = 1; $c -le $colCount; $c++) {
      $value = ''
      if ($c -le $row.Count) { $value = [string]$row[$c - 1] }
      $cellShape = $table.Cell($r, $c).Shape
      $cellShape.TextFrame.TextRange.Text = $value
      $cellShape.TextFrame.TextRange.Font.Size = 10
      $cellShape.TextFrame.MarginLeft = 4
      $cellShape.TextFrame.MarginRight = 4
      $cellShape.TextFrame.MarginTop = 3
      $cellShape.TextFrame.MarginBottom = 3
      if ($r -eq 1) {
        $cellShape.TextFrame.TextRange.Font.Bold = $true
        $cellShape.Fill.ForeColor.RGB = 0xEDEDED
      }
    }
  }
  return $slide
}

function Get-SlideText($Slide) {
  $parts = New-Object System.Collections.ArrayList
  foreach ($shape in @($Slide.Shapes)) {
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        [void]$parts.Add([string]$shape.TextFrame.TextRange.Text)
      }
    } catch {
    }
  }
  return ($parts -join "`n")
}

function Remove-DateSlides($Presentation, [string]$DateKey) {
  $removed = 0
  for ($i = $Presentation.Slides.Count; $i -ge 1; $i--) {
    $slide = $Presentation.Slides.Item($i)
    $text = Get-SlideText $slide
    if ($text -like "*$DateKey*") {
      $slide.Delete()
      $removed += 1
    }
  }
  return $removed
}

function Get-OneNoteTables($PageXml) {
  $tables = New-Object System.Collections.ArrayList
  foreach ($tableNode in @($PageXml.SelectNodes('//*[local-name()="Table"]'))) {
    $rows = New-Object System.Collections.ArrayList
    foreach ($rowNode in @($tableNode.SelectNodes('./*[local-name()="Row"]'))) {
      $cells = New-Object System.Collections.ArrayList
      foreach ($cellNode in @($rowNode.SelectNodes('./*[local-name()="Cell"]'))) {
        $cellTexts = @($cellNode.SelectNodes('.//*[local-name()="T"]') | ForEach-Object { $_.InnerText })
        [void]$cells.Add((Normalize-Text ($cellTexts -join "`n")))
      }
      if ($cells.Count -gt 0) { [void]$rows.Add(@($cells)) }
    }
    if ($rows.Count -gt 0) { [void]$tables.Add(@($rows)) }
  }
  return @($tables)
}

function Get-ImageExtensionFromBytes([byte[]]$Bytes, [string]$Fallback = 'png') {
  if ($Bytes.Length -ge 8 -and
      $Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x4E -and $Bytes[3] -eq 0x47) {
    return 'png'
  }
  if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xD8 -and $Bytes[2] -eq 0xFF) {
    return 'jpg'
  }
  if ($Bytes.Length -ge 6) {
    $asciiHead = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, [Math]::Min($Bytes.Length, 80))
    if ($asciiHead.Contains(' EMF')) { return 'emf' }
  }
  return $Fallback
}

try {
  $Device = Get-DeviceLabel
  $dateValue = [datetime]::Parse($Date)
  $datePatterns = Get-DatePatterns $dateValue
  $dateKey = $dateValue.ToString('yyyy-MM-dd')
  if ($Force) {
    $ReplaceDateSlides = $true
  }

  $pptxPath = Resolve-Path -LiteralPath $Pptx
  New-Item -ItemType Directory -Path $OutboxDir -Force | Out-Null
  $assetDir = Join-Path $OutboxDir ("assets_" + $dateValue.ToString('yyyyMMdd'))
  New-Item -ItemType Directory -Path $assetDir -Force | Out-Null
  $assetDirPath = (Resolve-Path -LiteralPath $assetDir).Path

  $statePath = Join-Path $OutboxDir 'onenote_experiment_ppt_state.json'
  $state = [ordered]@{ appended = @() }
  if (Test-Path -LiteralPath $statePath) {
    try {
      $loaded = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($loaded.appended) { $state.appended = @($loaded.appended) }
    } catch {
      $state = [ordered]@{ appended = @() }
    }
  }

  $one = New-Object -ComObject OneNote.Application
  [xml]$hierarchy = ''
  $one.GetHierarchy('', 4, [ref]$hierarchy)

  $notebook = @($hierarchy.DocumentElement.SelectNodes('//*[local-name()="Notebook"]') | Where-Object {
    $_.name -eq $NotebookName -or ($_.name -like "*$NotebookName*")
  } | Select-Object -First 1)
  if ($null -eq $notebook) {
    throw "Notebook not found: $NotebookName"
  }

  $sections = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
    $_.name -notlike '分類_*' -and $_.name -ne '削除されたページ'
  })
  if (-not [string]::IsNullOrWhiteSpace($SectionName)) {
    $sections = @($sections | Where-Object { $_.name -eq $SectionName -or (Get-SectionPath $_) -like "*$SectionName*" })
  }

  $matchedPages = New-Object System.Collections.ArrayList
  foreach ($section in $sections) {
    [xml]$sectionXml = ''
    try {
      $one.GetHierarchy($section.ID, 4, [ref]$sectionXml)
    } catch {
      continue
    }

    foreach ($page in @($sectionXml.DocumentElement.SelectNodes('.//*[local-name()="Page"]'))) {
      $pageXmlText = ''
      try {
        $one.GetPageContent($page.ID, [ref]$pageXmlText, 0)
      } catch {
        continue
      }

      [xml]$pageXml = $pageXmlText
      $texts = @($pageXml.SelectNodes('//*[local-name()="T"]') | ForEach-Object { $_.InnerText })
      $plain = Normalize-Text ($texts -join "`n")
      $haystack = (($page.name + "`n" + (Get-SectionPath $section) + "`n" + $plain))
      $isMatch = $false
      foreach ($pattern in $datePatterns) {
        if ($haystack.Contains($pattern)) {
          $isMatch = $true
          break
        }
      }
      if (-not $isMatch) { continue }

      $images = New-Object System.Collections.ArrayList
      $tables = Get-OneNoteTables $pageXml
      $imageIndex = 0
      foreach ($imageNode in @($pageXml.SelectNodes('//*[local-name()="Image"]'))) {
        $dataNode = @($imageNode.SelectNodes('.//*[local-name()="Data"]') | Select-Object -First 1)
        $imageB64 = ''
        if ($null -ne $dataNode -and -not [string]::IsNullOrWhiteSpace($dataNode.InnerText)) {
          $imageB64 = $dataNode.InnerText
        } else {
          $callbackNode = @($imageNode.SelectNodes('.//*[local-name()="CallbackID"]') | Select-Object -First 1)
          $callbackId = ''
          if ($null -ne $callbackNode -and $null -ne $callbackNode.callbackID) {
            $callbackId = [string]$callbackNode.callbackID
          }
          if (-not [string]::IsNullOrWhiteSpace($callbackId)) {
            try {
              $one.GetBinaryPageContent($page.ID, $callbackId, [ref]$imageB64)
            } catch {
              Write-Warning "Failed to read OneNote binary image content: $($page.name) / $($_.Exception.Message)"
              $imageB64 = ''
            }
          }
        }
        if ([string]::IsNullOrWhiteSpace($imageB64)) { continue }
        $imageBytes = [Convert]::FromBase64String($imageB64)
        $format = 'png'
        if ($imageNode.format) {
          $format = ([string]$imageNode.format).ToLowerInvariant()
        }
        if ($format -notin @('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff')) {
          $format = 'png'
        }
        $format = Get-ImageExtensionFromBytes $imageBytes $format
        $imageIndex += 1
        $safePage = ([string]$page.name -replace '[\\/:*?"<>|]', '_')
        if ($safePage.Length -gt 50) { $safePage = $safePage.Substring(0, 50) }
        $imagePath = Join-Path $assetDirPath ("{0}_{1}_{2}.{3}" -f $dateValue.ToString('yyyyMMdd'), $safePage, $imageIndex, $format)
        [System.IO.File]::WriteAllBytes($imagePath, $imageBytes)
        [void]$images.Add($imagePath)
      }

      $record = [ordered]@{
        page_id = [string]$page.ID
        title = [string]$page.name
        section = Get-SectionPath $section
        last_modified = [string]$page.lastModifiedTime
        text = $plain
        images = @($images)
        tables = @($tables)
      }
      [void]$matchedPages.Add($record)
    }
  }

  if ($matchedPages.Count -eq 0) {
    Write-Host "No matching OneNote pages found."
    Write-Host "Notebook: $NotebookName"
    Write-Host "Date: $dateKey"
    return
  }

  $fingerprintText = (($matchedPages | ForEach-Object { $_.page_id + $_.last_modified + $_.text + ($_.images -join ';') + (($_.tables | ConvertTo-Json -Depth 8 -Compress) -join '') }) -join "`n")
  $imageCount = ((@($matchedPages) | ForEach-Object { $_.images.Count } | Measure-Object -Sum).Sum)
  $tableCount = ((@($matchedPages) | ForEach-Object { $_.tables.Count } | Measure-Object -Sum).Sum)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($fingerprintText)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  $fingerprint = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
  if (-not $Force) {
    $existing = @($state.appended | Where-Object { $_.date -eq $dateKey -and $_.fingerprint -eq $fingerprint })
    if ($existing.Count -gt 0) {
      Write-Host "Skipped: same OneNote experiment content already appended."
      Write-Host "Date: $dateKey"
      Write-Host "Fingerprint: $fingerprint"
      return
    }
  }

  $snapshotPath = Join-Path $OutboxDir ("{0}_onenote_experiment_snapshot.json" -f $dateValue.ToString('yyyyMMdd'))
  $snapshot = [ordered]@{
    date = $dateKey
    notebook = $NotebookName
    section_filter = $SectionName
    page_count = $matchedPages.Count
    fingerprint = $fingerprint
    pages = @($matchedPages)
  }
  $snapshot | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $snapshotPath -Encoding UTF8

  if ($DryRun) {
    Write-Host "DryRun: yes"
    Write-Host "Notebook: $($notebook.name)"
    Write-Host "Section: $SectionName"
    Write-Host "Date: $dateKey"
    Write-Host "Pages: $($matchedPages.Count)"
    Write-Host "Images: $imageCount"
    Write-Host "Tables: $tableCount"
    Write-Host "Snapshot: $snapshotPath"
    return
  }

  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpoint.Visible = -1
  $presentation = $powerpoint.Presentations.Open($pptxPath.Path, $false, $false, $true)
  $removedSlides = 0
  if ($ReplaceDateSlides) {
    $removedSlides = Remove-DateSlides $presentation $dateKey
  }

  Add-TextSlide $presentation "$dateKey 実験ノート" "Device: $Device    Notebook: $($notebook.name)    Pages: $($matchedPages.Count)" "OneNoteの同日実験ノートから転記しました。`n`n対象ページ:`n$((@($matchedPages) | ForEach-Object { '- ' + $_.section + ' / ' + $_.title }) -join "`n")" | Out-Null

  foreach ($pageRecord in @($matchedPages)) {
    $chunks = @(Split-TextChunks $pageRecord.text 1100)
    for ($i = 0; $i -lt $chunks.Count; $i++) {
      $suffix = if ($chunks.Count -gt 1) { " ($($i + 1)/$($chunks.Count))" } else { '' }
      Add-TextSlide $presentation "$dateKey $($pageRecord.title)$suffix" "Section: $($pageRecord.section)    LastModified: $($pageRecord.last_modified)" $chunks[$i] | Out-Null
    }
    $imageNo = 0
    foreach ($imagePath in @($pageRecord.images)) {
      $imageNo += 1
      Add-ImageSlide $presentation "$dateKey $($pageRecord.title) image $imageNo" "Section: $($pageRecord.section)    Source: OneNote image" $imagePath | Out-Null
    }
    $tableNo = 0
    foreach ($tableRows in @($pageRecord.tables)) {
      $tableNo += 1
      Add-TableSlide $presentation "$dateKey $($pageRecord.title) table $tableNo" "Section: $($pageRecord.section)    Source: OneNote table" @($tableRows) | Out-Null
    }
  }

  $presentation.Save()
  $record = [ordered]@{
    appended_at = (Get-Date).ToString('s')
    date = $dateKey
    device = $Device
    pptx = $pptxPath.Path
    notebook = [string]$notebook.name
    section_filter = $SectionName
    page_count = $matchedPages.Count
    fingerprint = $fingerprint
    snapshot = $snapshotPath
    removed_slides = $removedSlides
    slide_count = $presentation.Slides.Count
  }
  $state.appended = @($state.appended) + @($record)
  $state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8

  Write-Host "PPTX: $($pptxPath.Path)"
  Write-Host "Date: $dateKey"
  Write-Host "Pages: $($matchedPages.Count)"
  Write-Host "Images: $imageCount"
  Write-Host "Tables: $tableCount"
  Write-Host "RemovedSlides: $removedSlides"
  Write-Host "SlideCount: $($presentation.Slides.Count)"
  Write-Host "Snapshot: $snapshotPath"
  Write-Host "State: $statePath"

  $presentation.Close()
  $powerpoint.Quit()
} finally {
  Pop-Location
}






