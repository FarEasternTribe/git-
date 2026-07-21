param(
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [string]$SectionName = '',
  [switch]$CommitBaseline
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONUTF8 = '1'
$root = Split-Path -Parent $PSScriptRoot
$sha = [System.Security.Cryptography.SHA256]::Create()
try {
  $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($PageTitle))
}
finally {
  $sha.Dispose()
}
$pageKey = ([System.BitConverter]::ToString($hashBytes) -replace '-', '').Substring(0, 16).ToLowerInvariant()
$relativePdf = "tmp\pdfs\${pageKey}_current.pdf"
$relativeOutputDir = "tmp\pdfs\${pageKey}_updates"
$relativeStateDir = ".agent_runtime\transcription_state\$pageKey"
$pdf = Join-Path $root $relativePdf
$outputDir = Join-Path $root $relativeOutputDir
$stateDir = Join-Path $root $relativeStateDir
$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) { $python = 'python' }

if (Test-Path -LiteralPath $outputDir) {
  Remove-Item -LiteralPath $outputDir -Recurse -Force
}
New-Item -ItemType Directory -Path (Split-Path -Parent $pdf) -Force | Out-Null

Push-Location $root
try {
  $exportArgs = @(
    '-ExecutionPolicy', 'Bypass',
    '-File', '.\tools\export_onenote_page_pdf.ps1',
    '-PageTitle', $PageTitle,
    '-Out', $relativePdf
  )
  if ($SectionName) { $exportArgs += @('-SectionName', $SectionName) }
  if ($PageTitle -match '^20\d{6}$') { $exportArgs += '-LatestForDate' }
  $exportOutput = @(& powershell @exportArgs)
  $exportExitCode = $LASTEXITCODE
  $exportOutput | ForEach-Object { Write-Output $_ }
  if ($exportExitCode -ne 0) { throw "OneNote PDF export failed: $exportExitCode" }

  if ($PageTitle -match '^20\d{6}$') {
    $resolvedTitleLine = @($exportOutput | Where-Object { [string]$_ -like 'ResolvedPageTitle=*' } | Select-Object -Last 1)
    $resolvedIdLine = @($exportOutput | Where-Object { [string]$_ -like 'ResolvedPageId=*' } | Select-Object -Last 1)
    $resolvedTitle = if ($resolvedTitleLine.Count) { ([string]$resolvedTitleLine[0]).Substring(18) } else { '' }
    $resolvedId = if ($resolvedIdLine.Count) { ([string]$resolvedIdLine[0]).Substring(15) } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($resolvedId) -and $resolvedTitle -cne $PageTitle) {
      $sourceHash = $sha = [System.Security.Cryptography.SHA256]::Create()
      try {
        $sourceHashBytes = $sha.ComputeHash(
          [System.Text.Encoding]::UTF8.GetBytes("$PageTitle`n$resolvedId")
        )
      }
      finally {
        $sha.Dispose()
      }
      $sourceKey = ([System.BitConverter]::ToString($sourceHashBytes) -replace '-', '').Substring(0, 16).ToLowerInvariant()
      $relativeStateDir = ".agent_runtime\transcription_state\$sourceKey"
      $stateDir = Join-Path $root $relativeStateDir
    }
  }

  $diffArgs = @(
    '.\tools\onenote_page_update_diff.py',
    '--pdf', $relativePdf,
    '--state-dir', $relativeStateDir,
    '--output-dir', $relativeOutputDir,
    '--dpi', '300'
  )
  if ($CommitBaseline) { $diffArgs += '--commit-baseline' }
  & $python @diffArgs
  if ($LASTEXITCODE -ne 0) { throw "OneNote update detection failed: $LASTEXITCODE" }
  Write-Output "UpdateImages=$outputDir"
  Write-Output "StateDir=$stateDir"
}
finally {
  if (Test-Path -LiteralPath $pdf) {
    Remove-Item -LiteralPath $pdf -Force
  }
  Pop-Location
}
