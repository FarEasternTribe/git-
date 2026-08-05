#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS Auto-Classifier
タイムライン入力 → Obsidian自動分類エンジン

入力フォーマット例:
16:42 Scholl反応開始
16:55 FeCl3を追加
17:20 溶液が黒色化
17:45 STM roomへ移動

自動分類先:
- 01_Daily （日常ログ）
- 03_Experiments （実験記録）
- 06_Meetings （会議）
- 07_Health （体調）
"""

import re
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple
import json

class TimelineClassifier:
    def __init__(self, vault_path: str):
        self.vault_path = Path(vault_path)

        # 分類ルール
        self.rules = {
            'experiment': {
                'keywords': ['反応', '合成', '触媒', 'FeCl3', 'スルホン酸', 'Scholl',
                           'AFM', 'STM', 'NMR', '収率', 'etch', 'pattern'],
                'destination': '03_Experiments',
                'template': 'TEMPLATE_Experiment'
            },
            'meeting': {
                'keywords': ['ゼミ', 'ミーティング', '会議', '打ち合わせ', 'meeting',
                           'M会', '発表', 'プレゼン'],
                'destination': '06_Meetings',
                'template': None
            },
            'health': {
                'keywords': ['睡眠', '体調', '痛み', '疲れ', '食事', '運動',
                           '体重', '帯状疱疹', '回復'],
                'destination': '07_Health',
                'template': None
            },
            'admin': {
                'keywords': ['メール', 'レポート', '採点', '学務', '書類',
                           '提出', '締切'],
                'destination': '08_Admin',
                'template': None
            },
            'default': {
                'keywords': [],
                'destination': '01_Daily',
                'template': None
            }
        }

    def classify_entry(self, text: str, timestamp: str = None) -> Tuple[str, str, Dict]:
        """
        単一エントリを分類

        Args:
            text: 入力テキスト
            timestamp: タイムスタンプ（HH:MM形式）

        Returns:
            (destination_folder, category, metadata)
        """
        if not timestamp:
            timestamp = datetime.now().strftime('%H:%M')

        # キーワードマッチング
        text_lower = text.lower()
        matched_category = 'default'
        max_match_count = 0

        for category, rule in self.rules.items():
            if category == 'default':
                continue

            match_count = sum(1 for kw in rule['keywords'] if kw.lower() in text_lower)
            if match_count > max_match_count:
                max_match_count = match_count
                matched_category = category

        rule = self.rules[matched_category]

        metadata = {
            'category': matched_category,
            'timestamp': timestamp,
            'original_text': text,
            'classified_at': datetime.now().isoformat(),
            'tags': self._generate_tags(matched_category, text)
        }

        return rule['destination'], matched_category, metadata

    def _generate_tags(self, category: str, text: str) -> List[str]:
        """テキストから自動タグ生成"""
        tags = []

        # カテゴリベースのタグ
        if category == 'experiment':
            tags.append('#synthesis')
            tags.append('#in-progress')
            if 'GNR' in text or 'gnr' in text:
                tags.append('#gnr')
            if 'Phosphorus' in text or 'phosphorus' in text:
                tags.append('#phosphorus')
        elif category == 'meeting':
            tags.append('#meeting')
        elif category == 'health':
            tags.append('#health')
            if '睡眠' in text:
                tags.append('#sleep')
            if '食事' in text:
                tags.append('#diet')
            if '体調' in text or '痛み' in text:
                tags.append('#recovery')
        elif category == 'admin':
            tags.append('#admin')
        else:
            tags.append('#work')

        return tags

    def create_obsidian_note(self, text: str, timestamp: str = None) -> Path:
        """
        Obsidianノートを自動作成

        Args:
            text: 入力テキスト
            timestamp: タイムスタンプ

        Returns:
            作成したファイルのパス
        """
        destination, category, metadata = self.classify_entry(text, timestamp)

        # ファイルパス
        now = datetime.now()
        filename = f"{now.strftime('%Y-%m-%d_%H%M%S')}_{category[:10]}.md"
        filepath = self.vault_path / destination / filename

        # フロントマター
        frontmatter = f"""---
type: timeline_entry
category: {category}
timestamp: {metadata['timestamp']}
created_at: {now.isoformat()}
tags: {json.dumps(metadata['tags'])}
---

# {now.strftime('%Y-%m-%d %H:%M')} — {category.upper()}

## 記録

{text}

## メモ

（自動分類）
"""

        # ファイル作成
        filepath.parent.mkdir(parents=True, exist_ok=True)
        filepath.write_text(frontmatter, encoding='utf-8')

        return filepath

    def process_timeline_batch(self, entries: List[Tuple[str, str]]) -> Dict:
        """
        複数エントリをバッチ処理

        Args:
            entries: [(timestamp, text), ...] のリスト

        Returns:
            処理結果辞書
        """
        results = {
            'processed': 0,
            'classified': {},
            'files_created': []
        }

        for timestamp, text in entries:
            try:
                filepath = self.create_obsidian_note(text, timestamp)
                destination, category, _ = self.classify_entry(text, timestamp)

                results['processed'] += 1
                results['classified'][category] = results['classified'].get(category, 0) + 1
                results['files_created'].append(str(filepath))

            except Exception as e:
                print(f"  ⚠️ エラー: {text[:30]}... - {e}")

        return results


def main():
    """メイン処理"""
    vault_path = r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳"

    classifier = TimelineClassifier(vault_path)

    # サンプル: タイムライン入力（実際にはFocus Desk PWAから来る）
    sample_entries = [
        ('09:44', 'AFM測定を開始。Tip交換後、チップ不良→J.Wangさん協力で解決'),
        ('10:15', 'AFM測定開始'),
        ('12:30', '昼食：パン'),
        ('14:00', '会議準備'),
        ('15:30', '睡眠不足のため体調注意'),
    ]

    print("🔄 タイムライン入力を自動分類中...")
    results = classifier.process_timeline_batch(sample_entries)

    print(f"\n✓ 処理完了: {results['processed']} エントリ")
    print("\n📊 分類結果:")
    for category, count in results['classified'].items():
        print(f"  {category}: {count}")

    print(f"\n📝 作成ファイル: {len(results['files_created'])}")
    for filepath in results['files_created'][:3]:
        print(f"  ✓ {Path(filepath).name}")


if __name__ == '__main__':
    main()
