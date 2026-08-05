#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS Daily Summary Generator
毎日の活動を自動でまとめる

出力:
- 今日の統計
- 優先度別タスク
- 進捗サマリー
"""

from datetime import datetime, timedelta
from pathlib import Path
import json
from typing import Dict
from research_os_search import ResearchOSSearch

class DailySummaryGenerator:
    def __init__(self, vault_path: str, db_path: str = "research_os.db"):
        self.vault_path = Path(vault_path)
        self.search_engine = ResearchOSSearch(vault_path, db_path)

    def generate_daily_summary(self, date: str = None) -> Dict:
        """
        指定日の日次要約を生成

        Args:
            date: ISO形式の日付（YYYY-MM-DD）。Noneなら今日

        Returns:
            要約辞書
        """
        if not date:
            date = datetime.now().strftime('%Y-%m-%d')

        # データベースから本日のデータを取得
        conn = self.search_engine.conn
        cursor = conn.cursor()

        # 本日のノート取得
        cursor.execute('''
        SELECT * FROM notes
        WHERE DATE(created_at) = ?
        ORDER BY created_at DESC
        ''', (date,))

        today_notes = [dict(row) for row in cursor.fetchall()]

        # タグ別集計
        cursor.execute('''
        SELECT tag, COUNT(*) as count FROM tags
        WHERE note_id IN (
            SELECT id FROM notes WHERE DATE(created_at) = ?
        )
        GROUP BY tag
        ORDER BY count DESC
        ''', (date,))

        tag_stats = {row['tag']: row['count'] for row in cursor.fetchall()}

        # カテゴリ別集計
        type_stats = {}
        for note in today_notes:
            doc_type = note.get('type', 'note')
            type_stats[doc_type] = type_stats.get(doc_type, 0) + 1

        summary = {
            'date': date,
            'total_notes': len(today_notes),
            'type_breakdown': type_stats,
            'tag_breakdown': tag_stats,
            'notes': today_notes,
            'priority_tasks': self._extract_priority_tasks(tag_stats),
            'progress_summary': self._generate_progress_summary(type_stats),
        }

        return summary

    def _extract_priority_tasks(self, tag_stats: Dict) -> Dict:
        """優先度別タスク抽出"""
        return {
            'urgent': tag_stats.get('#urgent', 0),
            'high': tag_stats.get('#high', 0),
            'normal': tag_stats.get('#normal', 0),
            'low': tag_stats.get('#low', 0),
        }

    def _generate_progress_summary(self, type_stats: Dict) -> str:
        """進捗サマリー生成"""
        if not type_stats:
            return "本日の記録がありません"

        parts = []
        for doc_type, count in sorted(type_stats.items(), key=lambda x: -x[1]):
            parts.append(f"{doc_type}×{count}")

        return " | ".join(parts)

    def export_summary_html(self, date: str = None) -> str:
        """
        HTML形式でエクスポート

        Returns:
            HTMLコード
        """
        summary = self.generate_daily_summary(date)

        html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Summary - {summary['date']}</title>
    <style>
        * {{ margin: 0; padding: 0; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                background: #f5f5f5; color: #333; padding: 20px; }}
        .container {{ max-width: 900px; margin: 0 auto; }}
        h1 {{ font-size: 28px; margin-bottom: 10px; }}
        .date {{ color: #999; font-size: 14px; margin-bottom: 20px; }}
        .stats {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }}
        .stat-box {{ background: white; padding: 15px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        .stat-box h3 {{ font-size: 12px; color: #999; text-transform: uppercase; margin-bottom: 10px; }}
        .stat-box .value {{ font-size: 32px; font-weight: bold; color: #0066cc; }}
        .breakdown {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        .breakdown h2 {{ font-size: 16px; margin-bottom: 15px; }}
        .breakdown-item {{ display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #eee; }}
        .breakdown-item:last-child {{ border-bottom: none; }}
        .tag {{ display: inline-block; background: #f0f0f0; padding: 4px 8px;
                border-radius: 4px; font-size: 12px; margin-right: 5px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Daily Summary</h1>
        <p class="date">{summary['date']}</p>

        <div class="stats">
            <div class="stat-box">
                <h3>Total Notes</h3>
                <div class="value">{summary['total_notes']}</div>
            </div>
            <div class="stat-box">
                <h3>Urgent</h3>
                <div class="value">{summary['priority_tasks']['urgent']}</div>
            </div>
            <div class="stat-box">
                <h3>High</h3>
                <div class="value">{summary['priority_tasks']['high']}</div>
            </div>
            <div class="stat-box">
                <h3>Normal</h3>
                <div class="value">{summary['priority_tasks']['normal']}</div>
            </div>
        </div>

        <div class="breakdown">
            <h2>Note Type Breakdown</h2>
            {''.join(f'<div class="breakdown-item"><span>{t}</span><span>{c}</span></div>'
                     for t, c in sorted(summary['type_breakdown'].items(), key=lambda x: -x[1]))}
        </div>
    </div>
</body>
</html>
"""
        return html


def main():
    """メイン処理"""
    vault_path = r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳"

    generator = DailySummaryGenerator(vault_path)

    # 本日の要約生成
    print("📊 本日の要約を生成中...")
    summary = generator.generate_daily_summary()

    print(f"\n✓ 生成完了: {summary['date']}")
    print(f"  総ノート数: {summary['total_notes']}")
    print(f"  進捗: {summary['progress_summary']}")

    print(f"\n🏷️ タグ統計（上位5件）:")
    for tag, count in sorted(summary['tag_breakdown'].items(), key=lambda x: -x[1])[:5]:
        print(f"  {tag}: {count}")

    generator.search_engine.close()


if __name__ == '__main__':
    main()
