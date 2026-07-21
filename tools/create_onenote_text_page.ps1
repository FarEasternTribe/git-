param(
  [Parameter(Mandatory=$true)][string]$SectionName,
  [Parameter(Mandatory=$true)][string]$PageTitle,
  [Parameter(Mandatory=$true)][string]$TextFile
)

$ErrorActionPreference = 'Stop'
$writer = Join-Path $PSScriptRoot 'create_onenote_text_image_page.ps1'
if (-not (Test-Path -LiteralPath $writer -PathType Leaf)) {
  throw "Shared OneNote writer not found: $writer"
}

& $writer -SectionName $SectionName -PageTitle $PageTitle -TextFile $TextFile
