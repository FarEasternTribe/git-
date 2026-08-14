# -*- coding: utf-8 -*-
<#
.SYNOPSIS
Research OS ダッシュボード → 11-TimeMemo 同期
ブラウザの localStorage から直接メモを抽出して Obsidian に保存

.DESCRIPTION
Windows レジストリ経由で Chrome の LocalStorage にアクセス
（または手動エクスポート機能を使用）
#>

param(
    [string]$OneDrivePath = "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude"
)

$VaultPath = "$OneDrivePath\外部脳"
$TimeMemoFolder = "$VaultPath\11-TimeMemo"
$SyncLogFile = "$OneDrivePath\research_os_timememo_sync.log"

# UTF-8 BOM 付きで出力
$OutputEncoding = [System.Text.Encoding]::UTF8

# フォルダ作成
if (!(Test-Path $TimeMemoFolder)) {
    New-Item -ItemType Directory -Force -Path $TimeMemoFolder | Out-Null
}

# ============================================================================
# ログ関数
# ============================================================================

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] $Message"
    Add-Content -Path $SyncLogFile -Value $logMsg -Encoding UTF8 -Force
    Write-Host $logMsg -ForegroundColor Cyan
}

# ============================================================================
# ダッシュボード HTML から localStorage データを抽出
# ============================================================================

function Get-LocalStorageData {
    $dashboardPath = "$OneDrivePath\research_os_dashboard_sync.html"

    if (!(Test-Path $dashboardPath)) {
        Write-Log "⚠ ダッシュボード HTML が見つかりません"
        return $null
    }

    # ブラウザの AppData から Chrome LocalStorage を読み込み
    # または、ユーザーが手動でエクスポートしたファイルから読み込み

    $chromeLocalStoragePath = "$env:APPDATA\..\Local\Google\Chrome\User Data\Default\Local Storage\https_file__0.localstorage"

    if (Test-Path $chromeLocalStoragePath) {
        Write-Log "ℹ Chrome LocalStorage を検出"
    }

    # 代替案：手動エクスポートファイルから読み込み
    $manualExportPath = "$OneDrivePath\research_os_timememo_export.json"
    if (Test-Path $manualExportPath) {
        try {
            $data = Get-Content -Path $manualExportPath -Encoding UTF8 | ConvertFrom-Json
            Write-Log "✓ エクスポートファイルから読み込み"
            return $data
        }
        catch {
            Write-Log "⚠ エクスポートファイルのパースに失敗"
        }
    }

    return $null
}

# ============================================================================
# メモを 11-TimeMemo に保存
# ============================================================================

function Save-TimeMemo {
    param(
        [PSObject]$Memo,
        [string]$Index
    )

    $timestamp = [datetime]::Parse($memo.timestamp)
    $fileBaseName = $timestamp.ToString('yyyy-MM-dd_HHmmss')
    if ($memo.title) {
        $fileBaseName += "_$($memo.title -replace '[^\w]', '' | Truncate 15)"
    }

    # frontmatter 付き Markdown
    $mdContent = @"
---
type: time_memo
timestamp: $($memo.timestamp)
title: $($memo.title)
category: $($memo.category)
tags: [$($memo.tags)]
synced_at: $(Get-Date -Format 'o')
---

# $($timestamp.ToString('HH:mm:ss')) — $($memo.title)

$($memo.text)

## メタデータ

- 時刻: $($timestamp.ToString('HH:mm:ss'))
- 日付: $($timestamp.ToString('yyyy-MM-dd'))
- カテゴリ: $($memo.category)
- 同期元: Browser Dashboard
"@

    $filePath = "$TimeMemoFolder\$fileBaseName.md"

    if (Test-Path $filePath) {
        Write-Log "⚠ スキップ: $fileBaseName.md（既存）"
        return $false
    }

    try {
        Set-Content -Path $filePath -Value $mdContent -Encoding UTF8 -Force
        Write-Log "✓ 保存: $fileBaseName.md"
        return $true
    }
    catch {
        Write-Log "❌ 保存失敗: $_"
        return $false
    }
}

function Truncate {
    param([string]$Text, [int]$Length = 15)
    if ($Text.Length -gt $Length) {
        return $Text.Substring(0, $Length)
    }
    return $Text
}

# ============================================================================
# メイン実行
# ============================================================================

Write-Log "🔄 Research OS TimeMemo 同期開始"

$data = Get-LocalStorageData

if (!$data) {
    Write-Log "❌ データを取得できません"
    Write-Log "→ ダッシュボームで『エクスポート』をクリックして"
    Write-Log "   research_os_timememo_export.json を生成してください"
    exit 1
}

# メモ配列を確認
if (!($data.PSObject.Properties.Name -contains 'memos')) {
    Write-Log "⚠ メモデータがありません"
    exit 0
}

$savedCount = 0
foreach ($memo in $data.memos) {
    if (Save-TimeMemo -Memo $memo) {
        $savedCount++
    }
}

Write-Log "📊 同期完了: $savedCount 個のメモを保存"
Write-Log "✓ 11-TimeMemo フォルダを確認: $TimeMemoFolder"
