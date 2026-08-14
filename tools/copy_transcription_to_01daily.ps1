param(
  [string]$Date = 'today',
  [string]$VaultRoot = ''
)

# 手書き文字起こしのミラー結果（外部脳\日常ログ\YYYY-MM-DD.md）を
# 外部脳\01_Daily\ にも複製する。
#
# mirror_transcription_to_vault.ps1 が作った正本をそのまま写すだけにしてある。
# 内容を作り直すと二重管理になり、片方だけ古くなるため。

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($VaultRoot)) {
  $VaultRoot = 'C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳'
}

# 日付の解釈
$target = $null
$s = $Date.Trim()
if ($s -eq 'today' -or $s -eq '今日')          { $target = (Get-Date).Date }
elseif ($s -eq 'yesterday' -or $s -eq '昨日')  { $target = (Get-Date).Date.AddDays(-1) }
elseif ($s -match '^(20\d{2})(\d{2})(\d{2})$') { $target = Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day ([int]$Matches[3]) }
elseif ($s -match '^(20\d{2})[-/](\d{1,2})[-/](\d{1,2})$') { $target = Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day ([int]$Matches[3]) }
elseif ($s -match '^(\d{1,2})[-/月](\d{1,2})日?$') { $target = Get-Date -Year (Get-Date).Year -Month ([int]$Matches[1]) -Day ([int]$Matches[2]) }
else { throw "日付を解釈できません: $Date" }

$stamp = $target.ToString('yyyy-MM-dd')

$srcDir = Join-Path $VaultRoot '日常ログ'
$dstDir = Join-Path $VaultRoot '01_Daily'
$src = Join-Path $srcDir "$stamp.md"
$dst = Join-Path $dstDir "$stamp.md"

if (-not (Test-Path -LiteralPath $src)) {
  Write-Output "Result=SOURCE_NOT_FOUND"
  Write-Output "Message=先に mirror_transcription_to_vault.ps1 を実行してください: $src"
  exit 2
}

if (-not (Test-Path -LiteralPath $dstDir)) {
  New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
}

# 既存があれば中身が同じか確認（同じなら書かない＝冪等）
$srcText = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
if (Test-Path -LiteralPath $dst) {
  $dstText = [System.IO.File]::ReadAllText($dst, [System.Text.Encoding]::UTF8)
  if ($dstText -eq $srcText) {
    Write-Output "Result=UNCHANGED"
    Write-Output "Path=$dst"
    exit 0
  }
  Write-Output "Note=既存ファイルを上書きします: $dst"
}

# OneDrive 上でも確実に書き込まれるよう WriteAllBytes を使う
$bytes = [System.Text.Encoding]::UTF8.GetBytes($srcText)
[System.IO.File]::WriteAllBytes($dst, (@(0xEF,0xBB,0xBF) + $bytes))

$info = Get-Item -LiteralPath $dst
Write-Output "Result=OK"
Write-Output "Path=$dst"
Write-Output "Bytes=$($info.Length)"
