param(
  [Parameter(Mandatory=$true)][string]$Date,
  [string]$NotebookName = 'FarEasternTribe',
  [string]$SectionName  = '手書き日誌',
  [string]$Out = ''
)

# FarEasternTribe > 手書き日誌 から、指定日の手書きページを特定して PDF に書き出す。
#
# このセクションのページはタブレットで書いたまま「無題のページ」になりがちで、
# タイトルだけでは日付を特定できない。そのため
#   1) タイトルに日付が含まれるページ
#   2) 作成日 or 最終更新日がその日のページ
# の両方を見て候補を出す。候補が1件に決まらないときは止めて一覧を出す。

$ErrorActionPreference = 'Stop'

# ---- 日付の解釈（8/5 / 8月5日 / 2026-08-05 / 20260805 を受ける） ----
function Resolve-TargetDate([string]$raw) {
  $s = $raw.Trim()
  $now = Get-Date

  if ($s -match '^(20\d{2})(\d{2})(\d{2})$') {
    return Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day ([int]$Matches[3]) -Hour 0 -Minute 0 -Second 0
  }
  if ($s -match '^(20\d{2})[-/\.](\d{1,2})[-/\.](\d{1,2})$') {
    return Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day ([int]$Matches[3]) -Hour 0 -Minute 0 -Second 0
  }
  if ($s -match '^(\d{1,2})[-/\.月](\d{1,2})日?$') {
    return Get-Date -Year $now.Year -Month ([int]$Matches[1]) -Day ([int]$Matches[2]) -Hour 0 -Minute 0 -Second 0
  }
  if ($s -eq 'today' -or $s -eq '今日') { return $now.Date }
  if ($s -eq 'yesterday' -or $s -eq '昨日') { return $now.Date.AddDays(-1) }

  throw "日付を解釈できません: $raw （例: 8/5, 8月5日, 2026-08-05, 20260805）"
}

$target = Resolve-TargetDate $Date
$stamp8 = $target.ToString('yyyyMMdd')
$stampDash = $target.ToString('yyyy-MM-dd')

Write-Output "TargetDate=$stampDash"

# ---- OneNote 階層を取得 ----
$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText

$notebook = @($hierarchy.SelectNodes('//*[local-name()="Notebook"]') |
  Where-Object { $_.name -eq $NotebookName })
if ($notebook.Count -ne 1) { throw "ノートブックが特定できません: $NotebookName (found $($notebook.Count))" }

$section = @($notebook[0].SelectNodes('.//*[local-name()="Section"]') |
  Where-Object { $_.name -eq $SectionName })
if ($section.Count -ne 1) { throw "セクションが特定できません: $NotebookName / $SectionName (found $($section.Count))" }

# ---- 同期してから読み直す（書いた直後の内容を取りこぼさないため） ----
$syncRequested = $false
try { $one.SyncHierarchy($notebook[0].ID); $syncRequested = $true; Start-Sleep -Seconds 3 }
catch { Write-Warning "SyncHierarchy failed: $($_.Exception.Message)" }

$sectionText = ''
$one.GetHierarchy($section[0].ID, 4, [ref]$sectionText, 2)
[xml]$sectionXml = $sectionText

# ---- 候補を集める ----
$candidates = @()
foreach ($p in @($sectionXml.SelectNodes('//*[local-name()="Page"]'))) {
  $title = [string]$p.name

  # 自動生成した文字起こしページは対象外（元ページだけを見る）
  if ($title -like '*文字起こし*') { continue }

  $created = $null; $modified = $null
  if ($p.dateTime)         { $created  = ([datetime]$p.dateTime).ToLocalTime() }
  if ($p.lastModifiedTime) { $modified = ([datetime]$p.lastModifiedTime).ToLocalTime() }

  $reasons = @()
  if ($title -match $stamp8 -or $title -match $stampDash) { $reasons += 'title' }
  if ($created  -ne $null -and $created.ToString('yyyy-MM-dd')  -eq $stampDash) { $reasons += 'created' }
  if ($modified -ne $null -and $modified.ToString('yyyy-MM-dd') -eq $stampDash) { $reasons += 'modified' }

  if ($reasons.Count -eq 0) { continue }

  $candidates += [pscustomobject]@{
    Title    = $title
    Id       = [string]$p.ID
    Created  = $created
    Modified = $modified
    Reason   = ($reasons -join '+')
  }
}

if ($candidates.Count -eq 0) {
  Write-Output "Result=NOT_FOUND"
  Write-Output "Message=$SectionName に $stampDash のページが見つかりません"
  exit 2
}

# タイトル一致を最優先、次に作成日一致、最後に更新日時が新しいもの
$ranked = $candidates | Sort-Object `
  @{ Expression = { if ($_.Reason -like '*title*') { 0 } else { 1 } } },
  @{ Expression = { if ($_.Reason -like '*created*') { 0 } else { 1 } } },
  @{ Expression = { $_.Modified }; Descending = $true }

Write-Output "CandidateCount=$($candidates.Count)"
$i = 0
foreach ($c in $ranked) {
  $i++
  $cs = ''; if ($c.Created)  { $cs = $c.Created.ToString('yyyy-MM-dd HH:mm') }
  $ms = ''; if ($c.Modified) { $ms = $c.Modified.ToString('yyyy-MM-dd HH:mm') }
  Write-Output ("Candidate{0}={1} | created={2} | modified={3} | match={4}" -f $i, $c.Title, $cs, $ms, $c.Reason)
}

$page = $ranked[0]

# 同点（同じ根拠で複数）なら自動で選ばず止める
$topReason = $page.Reason
$sameRank = @($ranked | Where-Object { $_.Reason -eq $topReason })
if ($sameRank.Count -gt 1) {
  Write-Output "Result=AMBIGUOUS"
  Write-Output "Message=候補が $($sameRank.Count) 件あります。どれを文字起こしするか指定してください。"
  exit 3
}

Write-Output "ResolvedPageTitle=$($page.Title)"
Write-Output "ResolvedPageId=$($page.Id)"
Write-Output "MatchReason=$($page.Reason)"
Write-Output "SyncRequested=$syncRequested"

# ---- PDF 書き出し ----
if ([string]::IsNullOrWhiteSpace($Out)) {
  $workspace = Split-Path -Parent $PSScriptRoot
  $Out = Join-Path $workspace ("tmp\pdfs\{0}_tegaki_nisshi.pdf" -f $stamp8)
}
$outDir = Split-Path -Parent $Out
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$absoluteOut = [System.IO.Path]::GetFullPath($Out)
if (Test-Path -LiteralPath $absoluteOut) { Remove-Item -LiteralPath $absoluteOut -Force }

# PublishFormat.pfPDF = 3
$one.Publish($page.Id, $absoluteOut, 3, '')
if (-not (Test-Path -LiteralPath $absoluteOut)) { throw "PDF was not created: $absoluteOut" }

Write-Output "Result=OK"
Write-Output "Pdf=$absoluteOut"
Write-Output "PdfBytes=$((Get-Item -LiteralPath $absoluteOut).Length)"
