param(
  [Parameter(Mandatory=$true)][string]$Date,
  [string]$NotebookName = 'FarEasternTribe',
  [string]$SectionName  = '手書き日誌',
  [string]$OutDir = ''
)

# FarEasternTribe > 手書き日誌 の指定日ページから、文字起こしの材料を取り出す。
#
# このPCには poppler(pdftoppm) が無く PDF をページ画像化できないため、PDF は経由しない。
# OneNote のページXMLから直接
#   - InkWord の recognizedText （OneNote自身の手書き認識。無料・オフライン・追加OCR不要）
#   - Image の base64 （写真/スクショ。PNGに書き出して後段で読む）
# を取り出す。

$ErrorActionPreference = 'Stop'

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
  if ($s -eq 'today' -or $s -eq '今日')     { return $now.Date }
  if ($s -eq 'yesterday' -or $s -eq '昨日') { return $now.Date.AddDays(-1) }
  throw "日付を解釈できません: $raw （例: 8/5, 8月5日, 2026-08-05, 20260805）"
}

$target    = Resolve-TargetDate $Date
$stamp8    = $target.ToString('yyyyMMdd')
$stampDash = $target.ToString('yyyy-MM-dd')
Write-Output "TargetDate=$stampDash"

$one = New-Object -ComObject OneNote.Application
$hierarchyText = ''
$one.GetHierarchy('', 4, [ref]$hierarchyText, 2)
[xml]$hierarchy = $hierarchyText

$notebook = @($hierarchy.SelectNodes('//*[local-name()="Notebook"]') | Where-Object { $_.name -eq $NotebookName })
if ($notebook.Count -ne 1) { throw "ノートブックが特定できません: $NotebookName (found $($notebook.Count))" }
$section = @($notebook[0].SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $SectionName })
if ($section.Count -ne 1) { throw "セクションが特定できません: $NotebookName / $SectionName (found $($section.Count))" }

# 書いた直後を取りこぼさないよう同期してから読む
$syncRequested = $false
try { $one.SyncHierarchy($notebook[0].ID); $syncRequested = $true; Start-Sleep -Seconds 3 }
catch { Write-Warning "SyncHierarchy failed: $($_.Exception.Message)" }
Write-Output "SyncRequested=$syncRequested"

$sectionText = ''
$one.GetHierarchy($section[0].ID, 4, [ref]$sectionText, 2)
[xml]$sectionXml = $sectionText

# ---- 日付でページを絞り込む ----
$candidates = @()
foreach ($p in @($sectionXml.SelectNodes('//*[local-name()="Page"]'))) {
  $title = [string]$p.name
  if ($title -like '*文字起こし*') { continue }   # 自動生成した出力ページは対象外

  $created = $null; $modified = $null
  if ($p.dateTime)         { $created  = ([datetime]$p.dateTime).ToLocalTime() }
  if ($p.lastModifiedTime) { $modified = ([datetime]$p.lastModifiedTime).ToLocalTime() }

  $reasons = @()
  if ($title -match $stamp8 -or $title -match $stampDash) { $reasons += 'title' }
  if ($created  -ne $null -and $created.ToString('yyyy-MM-dd')  -eq $stampDash) { $reasons += 'created' }
  if ($modified -ne $null -and $modified.ToString('yyyy-MM-dd') -eq $stampDash) { $reasons += 'modified' }
  if ($reasons.Count -eq 0) { continue }

  $candidates += [pscustomobject]@{
    Title = $title; Id = [string]$p.ID; Created = $created; Modified = $modified
    Reason = ($reasons -join '+')
  }
}

if ($candidates.Count -eq 0) {
  Write-Output "Result=NOT_FOUND"
  Write-Output "Message=$SectionName に $stampDash のページが見つかりません"
  exit 2
}

$ranked = $candidates | Sort-Object `
  @{ Expression = { if ($_.Reason -like '*title*')   { 0 } else { 1 } } },
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
$sameRank = @($ranked | Where-Object { $_.Reason -eq $page.Reason })
if ($sameRank.Count -gt 1) {
  Write-Output "Result=AMBIGUOUS"
  Write-Output "Message=候補が $($sameRank.Count) 件あります。どれを文字起こしするか指定してください。"
  exit 3
}

Write-Output "ResolvedPageTitle=$($page.Title)"
Write-Output "ResolvedPageId=$($page.Id)"
Write-Output "MatchReason=$($page.Reason)"

# ---- ページ本体を取得（piAll=4 でバイナリも含める） ----
$pageXmlText = ''
$one.GetPageContent($page.Id, [ref]$pageXmlText, 4)
[xml]$pageXml = $pageXmlText

if ([string]::IsNullOrWhiteSpace($OutDir)) {
  $workspace = Split-Path -Parent $PSScriptRoot
  $OutDir = Join-Path $workspace ("tmp\tegaki\{0}" -f $stamp8)
}
if (Test-Path -LiteralPath $OutDir) { Remove-Item -LiteralPath $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
Write-Output "OutDir=$OutDir"

# 位置（読み順の並べ替え用）。自分に無ければ祖先をたどる
function Get-Pos($node) {
  $n = $node
  while ($n -ne $null -and $n.NodeType -eq 'Element') {
    $p = $n.SelectSingleNode('*[local-name()="Position"]')
    if ($p -ne $null) {
      return [pscustomobject]@{ X = [double]$p.x; Y = [double]$p.y }
    }
    $n = $n.ParentNode
  }
  return [pscustomobject]@{ X = 0.0; Y = 0.0 }
}

$items = @()

# --- InkWord: OneNote が認識済みの手書きテキスト ---
foreach ($w in @($pageXml.SelectNodes('//*[local-name()="InkWord"]'))) {
  $pos = Get-Pos $w
  $items += [pscustomobject]@{ Kind='ink'; Text=[string]$w.recognizedText; X=$pos.X; Y=$pos.Y; Node=$w }
}

# --- 入力済みテキスト（手書きに混ざったタイプ入力） ---
foreach ($t in @($pageXml.SelectNodes('//*[local-name()="T"]'))) {
  $txt = [string]$t.InnerText
  if ([string]::IsNullOrWhiteSpace($txt)) { continue }
  $pos = Get-Pos $t
  $items += [pscustomobject]@{ Kind='text'; Text=$txt; X=$pos.X; Y=$pos.Y; Node=$t }
}

# --- Image: base64 を PNG に書き出す ---
$imgIndex = 0
foreach ($im in @($pageXml.SelectNodes('//*[local-name()="Image"]'))) {
  $dataNode = $im.SelectSingleNode('*[local-name()="Data"]')
  if ($dataNode -eq $null -or [string]::IsNullOrWhiteSpace($dataNode.InnerText)) { continue }
  $imgIndex++
  $file = Join-Path $OutDir ("image_{0:D2}.png" -f $imgIndex)
  try {
    [System.IO.File]::WriteAllBytes($file, [System.Convert]::FromBase64String($dataNode.InnerText))
  } catch {
    Write-Warning "画像の書き出しに失敗: $($_.Exception.Message)"
    continue
  }
  $pos = Get-Pos $im
  $items += [pscustomobject]@{ Kind='image'; Text=$file; X=$pos.X; Y=$pos.Y; Node=$im }
}

# --- InkDrawing: 図。認識テキストは無いので存在だけ知らせる ---
$inkDrawings = @($pageXml.SelectNodes('//*[local-name()="InkDrawing"]')).Count

# 読み順（上から下、同じ高さなら左から右）に並べる
$ordered = $items | Sort-Object @{Expression={[math]::Round($_.Y,0)}}, @{Expression={$_.X}}

$manifestLines = @()
foreach ($it in $ordered) {
  $manifestLines += ("{0}`t{1}`t{2}`t{3}" -f $it.Kind, [math]::Round($it.Y,0), [math]::Round($it.X,0), $it.Text)
}
$manifestPath = Join-Path $OutDir 'manifest.tsv'
[System.IO.File]::WriteAllText($manifestPath, ($manifestLines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))

# 認識済みテキストだけを読み順にまとめたもの
$inkText = ($ordered | Where-Object { $_.Kind -eq 'ink' -or $_.Kind -eq 'text' } | ForEach-Object { $_.Text }) -join "`r`n"
$inkPath = Join-Path $OutDir 'recognized.txt'
[System.IO.File]::WriteAllText($inkPath, $inkText, (New-Object System.Text.UTF8Encoding($true)))

$inkCount   = @($ordered | Where-Object { $_.Kind -eq 'ink' }).Count
$textCount  = @($ordered | Where-Object { $_.Kind -eq 'text' }).Count
$imageCount = @($ordered | Where-Object { $_.Kind -eq 'image' }).Count

Write-Output "InkWordCount=$inkCount"
Write-Output "TypedTextCount=$textCount"
Write-Output "ImageCount=$imageCount"
Write-Output "InkDrawingCount=$inkDrawings"
Write-Output "RecognizedTextFile=$inkPath"
Write-Output "Manifest=$manifestPath"

if ($inkCount -eq 0 -and $imageCount -eq 0 -and $textCount -eq 0) {
  Write-Output "Result=EMPTY"
  Write-Output "Message=このページには手書き・画像・テキストのいずれもありません（白紙）"
  exit 4
}

Write-Output "Result=OK"
