param(
  [int]$Days = 21,
  [string]$NotebookName = 'FarEasternTribe',
  [string[]]$SectionNames = @('実験'),
  [string]$OutDir = '.\agent_workspace\実験ノートAgent\experiment_board',
  [int]$TopN = 30
)

# 進行中実験トラッカー（試作）
# OneNote 2026実験/実験 の直近ページを読み、テキストから「どの段階か・次の一手」を
# キーワード推定し、要対応順のMarkdownボードを出力する。課金なし・ローカル処理・読み取りのみ。

$ErrorActionPreference = 'Stop'
$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $Workspace

function Get-PlainText($px) {
  $t = @($px.SelectNodes('//*[local-name()="T"]') | ForEach-Object { $_.InnerText })
  $s = ($t -join "`n")
  $s = [System.Net.WebUtility]::HtmlDecode($s)
  $s = $s -replace '(?is)<[^>]+>', ' '
  $s = $s -replace '\s+', ' '
  return $s.Trim()
}

# ステージ推定用キーワード（正規表現・大小無視）
$KW = [ordered]@{
  setup            = '仕掛け|反応開始|reaction start|入れた直後|撹拌|還流|reflux|stir|雰囲気下'
  monitor          = 'TLC|モニタ|チェック|check'
  workup           = '後処理|work[- ]?up|クエンチ|quench|分液|抽出|洗浄'
  purify           = 'カラム|column|再結晶|recryst|\d+\s*本目|回収|精製|フラッシュ'
  analysis         = 'NMR|LC[- ]?MS|\bMS\b|IR|測定|スペクトル|HRMS'
  analysis_pending = '取りに行く|明日.*(NMR|測定|データ)|(NMR|データ|測定).*(明日|後で|待|未)'
  done             = '収率|収量|yield|単離|isolated|\d+\s*mg.*得|完了'
}
# 実験ではない（PDF論文抽出・方法メモ）を除外
$excludePattern = 'PDF自動抽出|抽出文字数|根拠DOI|合成条件表（PDF|追加日時:.*PDF:'

$OutDirFull = (New-Item -ItemType Directory -Path $OutDir -Force).FullName

$one = New-Object -ComObject OneNote.Application
[xml]$h = ''
$one.GetHierarchy('', 4, [ref]$h)
$nb = @($h.DocumentElement.SelectNodes('//*[local-name()="Notebook"]') | Where-Object { $_.name -eq $NotebookName } | Select-Object -First 1)
if ($null -eq $nb) { throw "Notebook not found: $NotebookName" }

$since = (Get-Date).AddDays(-$Days)
$rows = New-Object System.Collections.ArrayList

foreach ($secName in $SectionNames) {
  $sec = @($nb.SelectNodes('.//*[local-name()="Section"]') | Where-Object { $_.name -eq $secName } | Select-Object -First 1)
  if ($null -eq $sec) { continue }
  [xml]$sx = ''
  $one.GetHierarchy($sec.ID, 4, [ref]$sx)
  foreach ($p in @($sx.DocumentElement.SelectNodes('.//*[local-name()="Page"]'))) {
    $modRaw = [string]$p.lastModifiedTime
    $mod = $null
    if ($modRaw) { try { $mod = [datetime]$modRaw } catch { $mod = $null } }
    if ($null -ne $mod -and $mod -lt $since) { continue }

    $xml = ''
    try { $one.GetPageContent($p.ID, [ref]$xml, 0) } catch { continue }
    [xml]$px = $xml
    $plain = Get-PlainText $px
    if ([string]::IsNullOrWhiteSpace($plain)) { continue }
    if ($plain -match $excludePattern) { continue }
    if ($plain.Length -lt 25) { continue }   # ほぼタイトルだけのページは除外

    $flags = [ordered]@{}
    foreach ($k in $KW.Keys) { $flags[$k] = [bool]([regex]::IsMatch($plain, $KW[$k], 'IgnoreCase')) }

    # ステージ・次の一手・優先度を推定
    if ($flags.done) {
      $stage = '完了（収率/単離あり）'; $next = '記録・NMR/収率の確認'; $prio = 3
    } elseif ($flags.analysis_pending -or ($flags.purify -and -not $flags.analysis)) {
      $stage = '分析待ち（精製済み）'; $next = 'NMR測定/データ回収'; $prio = 1
    } elseif ($flags.analysis) {
      $stage = '分析中'; $next = 'スペクトル解析・結果記録'; $prio = 2
    } elseif ($flags.workup) {
      $stage = '後処理済み'; $next = '精製（カラム/再結晶）へ'; $prio = 2
    } elseif ($flags.setup) {
      if ($flags.monitor) { $stage = '反応中（モニタ済み）'; $next = '後処理の判断' }
      else { $stage = '反応中/仕込み済み'; $next = 'TLCモニタ・後処理' }
      $prio = 1
    } else {
      $stage = '不明'; $next = '内容確認'; $prio = 2
    }

    $hit = @($KW.Keys | Where-Object { $flags[$_] }) -join ', '
    $snippet = if ($plain.Length -gt 90) { $plain.Substring(0, 90) + '…' } else { $plain }

    [void]$rows.Add([pscustomobject]@{
      Priority = $prio
      Section  = $secName
      Title    = [string]$p.name
      Modified = $mod
      Stage    = $stage
      Next     = $next
      Hits     = $hit
      Snippet  = $snippet
    })
  }
}

# 本文が薄くキーワードが取れないページ（古いサンプル作製メモ等）は本ボードから分離する。
$active = @($rows | Where-Object { -not ($_.Stage -eq '不明' -and [string]::IsNullOrWhiteSpace($_.Hits)) })
$lowinfo = @($rows | Where-Object { $_.Stage -eq '不明' -and [string]::IsNullOrWhiteSpace($_.Hits) })
$sorted = @($active | Sort-Object Priority, @{Expression='Modified';Descending=$true} | Select-Object -First $TopN)

$prioLabel = @{ 1 = '🔴要対応'; 2 = '🟡進行中'; 3 = '🟢完了近い' }
$now = Get-Date
$md = New-Object System.Collections.ArrayList
[void]$md.Add("# 進行中実験ボード（自動生成・試作）")
[void]$md.Add("")
[void]$md.Add("- 生成: $($now.ToString('yyyy-MM-dd HH:mm'))")
[void]$md.Add("- 対象: OneNote $NotebookName / $($SectionNames -join ', ') の直近 $Days 日（進行中 $($sorted.Count) 件）")
[void]$md.Add("- 段階・次の一手はテキストのキーワード推定（要確認）。課金なし・読み取りのみ。")
[void]$md.Add("")
[void]$md.Add("| 優先 | 推定ステージ | 次の一手 | 実験ページ | 最終更新 | 検出 |")
[void]$md.Add("|---|---|---|---|---|---|")
foreach ($r in $sorted) {
  $modStr = if ($r.Modified) { $r.Modified.ToString('MM-dd HH:mm') } else { '-' }
  [void]$md.Add(("| {0} | {1} | {2} | {3} | {4} | {5} |" -f $prioLabel[$r.Priority], $r.Stage, $r.Next, ($r.Title -replace '\|','/'), $modStr, ($r.Hits -replace '\|','/')))
}
[void]$md.Add("")
[void]$md.Add("## 詳細（先頭抜粋）")
foreach ($r in $sorted) {
  [void]$md.Add("- **[$($prioLabel[$r.Priority])] $($r.Title)** — $($r.Stage) → $($r.Next)")
  [void]$md.Add("  - $($r.Snippet)")
}
if ($lowinfo.Count -gt 0) {
  [void]$md.Add("")
  [void]$md.Add("## 判定不可（本文が薄い/タイトルのみ）")
  foreach ($r in @($lowinfo | Sort-Object @{Expression='Modified';Descending=$true})) {
    $modStr = if ($r.Modified) { $r.Modified.ToString('MM-dd') } else { '-' }
    [void]$md.Add("- $($r.Title)（$modStr）")
  }
}
$mdText = ($md -join "`n")

$stamp = $now.ToString('yyyyMMdd_HHmmss')
$outPath = Join-Path $OutDirFull "${stamp}_experiment_board.md"
$latestPath = Join-Path $OutDirFull 'latest_experiment_board.md'
Set-Content -LiteralPath $outPath -Value $mdText -Encoding UTF8
Set-Content -LiteralPath $latestPath -Value $mdText -Encoding UTF8

Write-Host "Board: $outPath"
Write-Host "Latest: $latestPath"
Write-Host "Rows: $($sorted.Count)"
Write-Host ""
Write-Host $mdText

Pop-Location
