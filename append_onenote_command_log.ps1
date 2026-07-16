param(
  [string]$NotebookName = 'OpenAI_Agent1',
  [string]$SectionName = '命令したLog_Claude',
  [string]$Device = '',
  [string]$Actor = 'Claude',
  [Parameter(Mandatory = $true)]
  [string]$Summary,
  [string[]]$Actions = @(),
  [string[]]$Files = @(),
  [string[]]$Verification = @(),
  [string[]]$RequiredOnOtherDevice = @(),
  [string[]]$NextSteps = @()
)

$ErrorActionPreference = 'Stop'

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutboxDir = Join-Path $Workspace 'agent_workspace\司令塔Agent\command_log_outbox'
New-Item -ItemType Directory -Path $OutboxDir -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($Device)) {
  $computerName = [string]$env:COMPUTERNAME
  if ($computerName.ToUpperInvariant().Contains('LENOVO')) {
    $Device = 'Lenovo'
  } elseif ($computerName.ToUpperInvariant().Contains('DESKTOP')) {
    $Device = 'Desktop'
  } elseif (-not [string]::IsNullOrWhiteSpace($computerName)) {
    $Device = $computerName
  } else {
    $Device = 'UnknownPC'
  }
}

function Escape-OneNoteXml([string]$Text) {
  if ($null -eq $Text) { return '' }
  return [System.Security.SecurityElement]::Escape($Text)
}

function New-ListMarkdown([string]$Title, [string[]]$Items) {
  $lines = @("## $Title", '')
  $expanded = @()
  foreach ($item in @($Items)) {
    if ($null -eq $item) { continue }
    foreach ($part in ([string]$item -split ';')) {
      if (-not [string]::IsNullOrWhiteSpace($part)) {
        $expanded += $part.Trim()
      }
    }
  }
  if ($expanded.Count -eq 0) {
    $lines += '- なし'
  } else {
    foreach ($item in $expanded) {
      $lines += "- $item"
    }
  }
  return $lines -join "`n"
}

$now = Get-Date
$stamp = $now.ToString('yyyyMMdd_HHmmss')
$title = "[$Device] " + $now.ToString('yyyy-MM-dd HH:mm') + " " + $Summary
$safeTitle = ($title -replace '[\\/:*?"<>|]', '_')
# Windows MAX_PATH(260) ガード: OneDrive の深いベースパスでも
# <OutboxDir>\<stamp>_<title>.md が上限を超えないようファイル名用の題名長を抑える。
# (stamp は 'yyyyMMdd_HHmmss' 固定15文字 + 区切り2文字 + '.md' 3文字 = 20文字を予約。
#  OneNoteページ側のタイトルは $title のまま全文を使うのでここでの短縮はファイル名だけに影響)
$maxTitleLen = 250 - $OutboxDir.Length - 20
if ($maxTitleLen -gt 110) { $maxTitleLen = 110 }
if ($maxTitleLen -lt 8) { $maxTitleLen = 8 }
if ($safeTitle.Length -gt $maxTitleLen) {
  $safeTitle = $safeTitle.Substring(0, $maxTitleLen)
}

$markdown = @(
  "# $title",
  '',
  "- Device: $Device",
  "- Actor: $Actor",
  "- Timestamp: $($now.ToString('s'))",
  "- Notebook: $NotebookName",
  "- Section: $SectionName",
  '',
  "## Summary",
  '',
  $Summary,
  '',
  (New-ListMarkdown 'Actions' $Actions),
  '',
  (New-ListMarkdown 'Files' $Files),
  '',
  (New-ListMarkdown 'Verification' $Verification),
  '',
  (New-ListMarkdown 'Required On Other Device' $RequiredOnOtherDevice),
  '',
  (New-ListMarkdown 'Next steps' $NextSteps)
) -join "`n"

$localPath = Join-Path $OutboxDir "${stamp}_${safeTitle}.md"
Set-Content -LiteralPath $localPath -Value $markdown -Encoding UTF8

try {
  $one = New-Object -ComObject OneNote.Application
  [xml]$hierarchy = ''
  $one.GetHierarchy('', 4, [ref]$hierarchy)

  $notebook = @($hierarchy.DocumentElement.SelectNodes('//*[local-name()="Notebook"]') | Where-Object {
    $_.name -eq $NotebookName
  } | Select-Object -First 1)
  if ($null -eq $notebook) {
    throw "Notebook not found: $NotebookName"
  }

  $section = @($notebook.SelectNodes('.//*[local-name()="Section"]') | Where-Object {
    $_.name -eq $SectionName
  } | Select-Object -First 1)
  if ($null -eq $section) {
    # 初回実行時はClaude専用セクションがまだ無いので作成する（cfSection = 3）
    $newSectionId = ''
    $one.OpenHierarchy($notebook.path + $SectionName + '.one', '', [ref]$newSectionId, 3)
    if ([string]::IsNullOrWhiteSpace($newSectionId)) {
      throw "Section not found or could not be created: $NotebookName / $SectionName"
    }
    $section = [pscustomobject]@{ ID = $newSectionId }
  }

  $pageId = ''
  $one.CreateNewPage($section.ID, [ref]$pageId, 0)

  $bodyLines = @()
  foreach ($line in $markdown -split "`n") {
    $bodyLines += '<one:OE><one:T><![CDATA[' + $line + ']]></one:T></one:OE>'
  }
  $bodyXml = $bodyLines -join "`n"
  $pageXml = @"
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="$pageId">
  <one:Title>
    <one:OE>
      <one:T><![CDATA[$title]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36" y="86" z="0" />
    <one:Size width="900" height="600" />
    <one:OEChildren>
$bodyXml
    </one:OEChildren>
  </one:Outline>
</one:Page>
"@
  $one.UpdatePageContent($pageXml)
  Write-Host "OneNote: ok"
  Write-Host "PageId: $pageId"
} catch {
  Write-Warning "OneNote write failed: $($_.Exception.Message)"
  Write-Host "OneNote: failed"
}

Write-Host "LocalLog: $localPath"




