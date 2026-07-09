param(
  [string]$Pptx = '.\Experiment.pptx',
  [string]$SourceMarkdown = '',
  [string]$Title = '',
  [string]$ExperimentDate = '',
  [string]$Sample = '',
  [string]$Objective = '',
  [string]$Procedure = '',
  [string]$Conditions = '',
  [string]$Observations = '',
  [string]$Results = '',
  [string]$NextActions = '',
  [string]$Device = 'Desktop',
  [string]$OutboxDir = '.\agent_workspace\実験ノートAgent\logs',
  [switch]$Force
)

$ErrorActionPreference = 'Stop'

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Workspace

function Normalize-Text([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return '記載なし' }
  return (($Text -replace "`r", '') -replace "`n{3,}", "`n`n").Trim()
}

function Truncate-Text([string]$Text, [int]$MaxChars) {
  $value = Normalize-Text $Text
  if ($value.Length -le $MaxChars) { return $value }
  return $value.Substring(0, $MaxChars - 1) + '…'
}

function Get-MarkdownSection([string]$Markdown, [string[]]$Names) {
  foreach ($name in $Names) {
    $pattern = "(?ms)^##\s*(?:\d+[.．]\s*)?$([regex]::Escape($name))\s*`n(.*?)(?=^##\s|\z)"
    $match = [regex]::Match($Markdown, $pattern)
    if ($match.Success) {
      return $match.Groups[2].Value.Trim()
    }
  }
  return ''
}

try {
  $pptxPath = Resolve-Path -LiteralPath $Pptx
  if ([string]::IsNullOrWhiteSpace($ExperimentDate)) {
    $ExperimentDate = (Get-Date).ToString('yyyy-MM-dd')
  }

  $sourceText = ''
  $sourcePath = ''
  $sourceHash = ''
  if (-not [string]::IsNullOrWhiteSpace($SourceMarkdown)) {
    $sourcePath = (Resolve-Path -LiteralPath $SourceMarkdown).Path
    $sourceText = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
    $sourceBytes = [System.Text.Encoding]::UTF8.GetBytes($sourceText)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $sourceHash = [System.BitConverter]::ToString($sha.ComputeHash($sourceBytes)).Replace('-', '').ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($Title)) {
      $h1 = [regex]::Match($sourceText, '(?m)^#\s+(.+)$')
      if ($h1.Success) {
        $Title = $h1.Groups[1].Value.Trim()
      }
    }
    if ([string]::IsNullOrWhiteSpace($Objective)) {
      $Objective = Get-MarkdownSection $sourceText @('研究', '実験', '今日行った研究・実験・解析')
    }
    if ([string]::IsNullOrWhiteSpace($Procedure)) {
      $Procedure = Get-MarkdownSection $sourceText @('研究', '実験')
    }
    if ([string]::IsNullOrWhiteSpace($Observations)) {
      $Observations = Get-MarkdownSection $sourceText @('得られた結果', '結果', '観察', '研究')
    }
    if ([string]::IsNullOrWhiteSpace($NextActions)) {
      $NextActions = Get-MarkdownSection $sourceText @('明日以降のTODO', 'TODO候補', '今後確認したいこと', '次に確認すべきこと')
    }
  }

  if ([string]::IsNullOrWhiteSpace($Title)) {
    $Title = "$ExperimentDate 実験ノート"
  }

  New-Item -ItemType Directory -Path $OutboxDir -Force | Out-Null
  $statePath = Join-Path $OutboxDir 'experiment_note_state.json'
  $state = [ordered]@{ appended = @() }
  if (Test-Path -LiteralPath $statePath) {
    try {
      $loaded = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($loaded.appended) {
        $state.appended = @($loaded.appended)
      }
    } catch {
      $state = [ordered]@{ appended = @() }
    }
  }

  if ($sourceHash -and -not $Force) {
    $existing = @($state.appended | Where-Object { $_.source_hash -eq $sourceHash -and $_.source_path -eq $sourcePath })
    if ($existing.Count -gt 0) {
      Write-Host "Skipped: source already appended"
      Write-Host "Source: $sourcePath"
      Write-Host "Hash: $sourceHash"
      return
    }
  }

  $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
  $safeTitle = ($Title -replace '[\\/:*?"<>|]', '_')
  if ($safeTitle.Length -gt 80) { $safeTitle = $safeTitle.Substring(0, 80) }
  $markdownLog = Join-Path $OutboxDir "${stamp}_${safeTitle}.md"

  $noteMarkdown = @(
    "# $Title",
    '',
    "- Device: $Device",
    "- Date: $ExperimentDate",
    "- PPTX: $($pptxPath.Path)",
    "- Source: $sourcePath",
    '',
    "## Sample",
    (Normalize-Text $Sample),
    '',
    "## Objective",
    (Normalize-Text $Objective),
    '',
    "## Procedure",
    (Normalize-Text $Procedure),
    '',
    "## Conditions",
    (Normalize-Text $Conditions),
    '',
    "## Observations",
    (Normalize-Text $Observations),
    '',
    "## Results",
    (Normalize-Text $Results),
    '',
    "## Next actions",
    (Normalize-Text $NextActions)
  ) -join "`n"
  Set-Content -LiteralPath $markdownLog -Value $noteMarkdown -Encoding UTF8

  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpoint.Visible = -1
  $presentation = $powerpoint.Presentations.Open($pptxPath.Path, $false, $false, $true)

  $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
  $slideWidth = $presentation.PageSetup.SlideWidth
  $slideHeight = $presentation.PageSetup.SlideHeight

  $margin = 32
  $titleBox = $slide.Shapes.AddTextbox(1, $margin, 20, $slideWidth - ($margin * 2), 46)
  $titleBox.TextFrame.TextRange.Text = $Title
  $titleBox.TextFrame.TextRange.Font.Size = 24
  $titleBox.TextFrame.TextRange.Font.Bold = $true

  $metaBox = $slide.Shapes.AddTextbox(1, $margin, 66, $slideWidth - ($margin * 2), 30)
  $metaBox.TextFrame.TextRange.Text = "Date: $ExperimentDate    Device: $Device    Source: " + ($(if ($sourcePath) { Split-Path $sourcePath -Leaf } else { 'manual' }))
  $metaBox.TextFrame.TextRange.Font.Size = 10
  $metaBox.TextFrame.TextRange.Font.Color.RGB = 0x666666

  $left = $margin
  $top = 108
  $colGap = 20
  $colWidth = ($slideWidth - ($margin * 2) - $colGap) / 2
  $boxHeight = ($slideHeight - $top - 42) / 3

  $blocks = @(
    @{ Label = 'Sample'; Text = $Sample },
    @{ Label = 'Objective'; Text = $Objective },
    @{ Label = 'Procedure'; Text = $Procedure },
    @{ Label = 'Conditions'; Text = $Conditions },
    @{ Label = 'Observations / Results'; Text = (($Observations, $Results | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n") },
    @{ Label = 'Next actions'; Text = $NextActions }
  )

  for ($i = 0; $i -lt $blocks.Count; $i++) {
    $row = [math]::Floor($i / 2)
    $col = $i % 2
    $x = $left + ($col * ($colWidth + $colGap))
    $y = $top + ($row * ($boxHeight + 10))
    $shape = $slide.Shapes.AddTextbox(1, $x, $y, $colWidth, $boxHeight)
    $text = $blocks[$i].Label + "`n" + (Truncate-Text $blocks[$i].Text 520)
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Size = 11
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 2
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 6
    $shape.Line.ForeColor.RGB = 0xD0D0D0
    $shape.Fill.ForeColor.RGB = 0xFAFAFA
    $shape.TextFrame.TextRange.Characters(1, $blocks[$i].Label.Length).Font.Bold = $true
    $shape.TextFrame.TextRange.Characters(1, $blocks[$i].Label.Length).Font.Size = 12
  }

  $presentation.Save()
  $record = [ordered]@{
    appended_at = (Get-Date).ToString('s')
    title = $Title
    device = $Device
    pptx = $pptxPath.Path
    source_path = $sourcePath
    source_hash = $sourceHash
    markdown_log = $markdownLog
    slide_count = $presentation.Slides.Count
  }
  $state.appended = @($state.appended) + @($record)
  $state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8
  Write-Host "PPTX: $($pptxPath.Path)"
  Write-Host "SlideCount: $($presentation.Slides.Count)"
  Write-Host "MarkdownLog: $markdownLog"
  Write-Host "State: $statePath"
  $presentation.Close()
  $powerpoint.Quit()
} finally {
  Pop-Location
}


