# -*- coding: utf-8 -*-
<#
.SYNOPSIS
Claude の作業ログを OneNote に記録

.USAGE
.\record_claude_work_log.ps1 -TextFile "C:\path\to\log.txt" -PageTitle "[Desktop] 作業内容"
#>

param(
    [Parameter(Mandatory=$true)][string]$TextFile,
    [Parameter(Mandatory=$true)][string]$PageTitle,
    [string]$NotebookName = 'FarEasternTribe',
    [string]$SectionName = '命令したLog_Claude'
)

$ErrorActionPreference = 'Stop'

try {
    # テキストファイルを読み込み
    $textContent = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)

    # OneNote COM オブジェクトを初期化
    $oneNote = New-Object -ComObject OneNote.Application

    # ノートブック一覧を取得
    $hierarchyXml = ''
    $oneNote.GetHierarchy('', 4, [ref]$hierarchyXml, 2)

    [xml]$hierarchyDoc = $hierarchyXml
    $ns = New-Object System.Xml.XmlNamespaceManager($hierarchyDoc.NameTable)
    $ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')

    # ノートブックを取得
    $notebook = $hierarchyDoc.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $ns)
    if (-not $notebook) {
        throw "Notebook not found: $NotebookName"
    }
    $notebookId = $notebook.GetAttribute('ID')

    # セクション一覧を取得
    $sectionXml = ''
    $oneNote.GetHierarchy($notebookId, 1, [ref]$sectionXml, 2)

    [xml]$sectionDoc = $sectionXml
    $section = $sectionDoc.SelectSingleNode("//one:Section[@name='$SectionName']", $ns)
    if (-not $section) {
        throw "Section not found: $SectionName"
    }
    $sectionId = $section.GetAttribute('ID')

    # 新しいページを作成
    $pageId = ''
    $oneNote.CreateNewPage($sectionId, [ref]$pageId, 1)

    # ページコンテンツを構築（XML形式）
    $namespace = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

    # テキストを複数の段落に分割
    $paragraphs = $textContent -split "`n" | Where-Object { $_.Trim() }

    $outlineXml = '<one:Outline xmlns:one="' + $namespace + '">'
    $outlineXml += '<one:OEChildren>'

    # ページタイトル
    $outlineXml += '<one:OE><one:T><![CDATA[' + $PageTitle + ']]></one:T></one:OE>'

    # テキスト行を追加
    foreach ($line in $paragraphs) {
        $cleanedLine = [System.Security.SecurityElement]::Escape($line)
        $outlineXml += '<one:OE><one:T><![CDATA[' + $line + ']]></one:T></one:OE>'
    }

    $outlineXml += '</one:OEChildren>'
    $outlineXml += '</one:Outline>'

    # ページ更新用XML
    $pageContent = '<?xml version="1.0" encoding="UTF-8"?>'
    $pageContent += '<one:Page xmlns:one="' + $namespace + '" '
    $pageContent += 'formatVersion="2.0" '
    $pageContent += 'creationTime="' + (Get-Date -Format 'o') + '" '
    $pageContent += 'lastModifiedTime="' + (Get-Date -Format 'o') + '" '
    $pageContent += 'ID="' + $pageId + '">'
    $pageContent += '<one:Title><one:OEChildren><one:OE><one:T><![CDATA[' + $PageTitle + ']]></one:T></one:OE></one:OEChildren></one:Title>'
    $pageContent += $outlineXml
    $pageContent += '</one:Page>'

    # ページコンテンツを更新
    $oneNote.UpdatePageContent($pageContent)

    Write-Host "✓ OneNote に記録完了"
    Write-Host "  ノートブック: $NotebookName"
    Write-Host "  セクション: $SectionName"
    Write-Host "  ページ: $PageTitle"

} catch {
    Write-Host "❌ エラー: $_" -ForegroundColor Red
    exit 1
}
