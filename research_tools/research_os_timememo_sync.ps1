# -*- coding: utf-8 -*-
<#
.SYNOPSIS
Research OS TimeMemo Sync
ブラウザ LocalStorage から 11-TimeMemo フォルダに自動同期

.DESCRIPTION
- ダッシュボードのメモデータを監視
- JSON ファイルから読み込み
- Obsidian 11-TimeMemo フォルダに .md ファイル生成
- 定期実行（タスク スケジューラ）
#>

param(
    [string]$OneDrivePath = "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude",
    [int]$MaxRetries = 3
)

# ============================================================================
# 初期化
# ============================================================================

$VaultPath = "$OneDrivePath\外部脳"
$TimeMemoFolder = "$VaultPath\11-TimeMemo"
$SyncDataFile = "$OneDrivePath\research_os_timememo_queue.json"
$LogFile = "$OneDrivePath\research_os_sync.log"

# UTF-8 BOM 付きで出力
$OutputEncoding = [System.Text.Encoding]::UTF8

# フォルダ作成
if (!(Test-Path $TimeMemoFolder)) {
    New-Item -ItemType Directory -Force -Path $TimeMemoFolder | Out-Null
    Write-Host "✓ 11-TimeMemo フォルダ作成"
}

# ============================================================================
# ログ関数
# ============================================================================

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] $Message"
    Add-Content -Path $LogFile -Value $logMsg -Encoding UTF8 -Force
    Write-Host $logMsg
}

# ============================================================================
# メイン処理
# ============================================================================

function Sync-TimeMemos {
    try {
        # 既に存在する .md ファイルをスキャン（重複チェック）
        $existingFiles = Get-ChildItem -Path $TimeMemoFolder -Filter "*.md" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty BaseName

        Write-Log "📝 11-TimeMemo 同期開始（既存ファイル数: $($existingFiles.Count)）"

        # ダッシュボーム HTML ファイルから LocalStorage を抽出
        $dashboardPath = "$OneDrivePath\research_os_dashboard_sync.html"

        if (!(Test-Path $dashboardPath)) {
            Write-Log "⚠ ダッシュボード HTML が見つかりません"
            return
        }

        # JSON キューファイルを確認
        $savedCount = 0
        if (Test-Path $SyncDataFile) {
            try {
                $queueData = Get-Content -Path $SyncDataFile -Encoding UTF8 | ConvertFrom-Json

                if ($queueData -and $queueData.PSObject.Properties.Name -contains 'memos') {
                    $memos = $queueData.memos

                    foreach ($memo in $memos) {
                        # ファイル名生成（重複回避）
                        $timestamp = [datetime]::Parse($memo.timestamp)
                        $fileBaseName = "$($timestamp.ToString('yyyy-MM-dd_HHmmss'))_$($memo.title -replace '[^\w\s]', '' | Truncate 20)"

                        if ($existingFiles -contains $fileBaseName) {
                            Write-Log "ℹ スキップ: $fileBaseName（既存）"
                            continue
                        }

                        # frontmatter 付き Markdown 生成
                        $mdContent = @"
---
type: time_memo
timestamp: $($memo.timestamp)
title: $($memo.title)
category: $($memo.category -or 'default')
tags: [$($memo.tags -split '\s+' | ForEach-Object { "'$_'" } | Join-String -Separator ', ')]
synced_at: $(Get-Date -Format 'o')
---

# $($timestamp.ToString('HH:mm:ss')) — $($memo.title)

$($memo.text)

## メタデータ

- 時刻: $($timestamp.ToString('HH:mm:ss'))
- 日付: $($timestamp.ToString('yyyy-MM-dd'))
- 同期: ✓ Desktop from Browser
"@

                        # ファイル保存
                        $filePath = "$TimeMemoFolder\$fileBaseName.md"
                        Set-Content -Path $filePath -Value $mdContent -Encoding UTF8 -Force

                        Write-Log "✓ 保存: $fileBaseName.md"
                        $savedCount++
                    }

                    # キューファイルを削除（処理完了）
                    Remove-Item -Path $SyncDataFile -Force -ErrorAction SilentlyContinue
                    Write-Log "✓ キューファイル削除"
                }
            }
            catch {
                Write-Log "⚠ JSON パースエラー: $_"
            }
        }

        # 完了ログ
        if ($savedCount -gt 0) {
            Write-Log "📊 同期完了: $savedCount メモを保存"
        } else {
            Write-Log "ℹ 新しいメモなし"
        }

    }
    catch {
        Write-Log "❌ エラー: $_"
    }
}

# Truncate helper
filter Truncate {
    param([int]$Length = 20)
    if ($_.Length -gt $Length) {
        $_.Substring(0, $Length)
    } else {
        $_
    }
}

# ============================================================================
# 実行
# ============================================================================

Sync-TimeMemos
Write-Log "---"
