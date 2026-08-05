#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS Search Engine
小島さん専用のObsidian全文検索エンジン

機能：
- Markdownファイルのインデックス化
- メタデータ（タグ、作成日時）抽出
- 意味検索 + 全文検索
- 日次要約生成
"""

import sqlite3
import os
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple
import json
import yaml

class ResearchOSSearch:
    def __init__(self, vault_path: str, db_path: str = "research_os.db"):
        """
        初期化

        Args:
            vault_path: Obsidianボルトのパス
            db_path: SQLiteデータベースのパス
        """
        self.vault_path = Path(vault_path)
        self.db_path = Path(db_path)
        self.conn = None
        self._init_database()

    def _init_database(self):
        """SQLiteデータベースを初期化"""
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        cursor = self.conn.cursor()

        # ノートテーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filepath TEXT UNIQUE NOT NULL,
            title TEXT,
            content TEXT,
            type TEXT,
            created_at TEXT,
            modified_at TEXT,
            tags TEXT,
            metadata TEXT,
            indexed_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        ''')

        # 検索インデックステーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS search_index (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            note_id INTEGER,
            word TEXT,
            frequency INTEGER,
            FOREIGN KEY (note_id) REFERENCES notes(id)
        )
        ''')

        # タグテーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS tags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            note_id INTEGER,
            tag TEXT,
            FOREIGN KEY (note_id) REFERENCES notes(id)
        )
        ''')

        self.conn.commit()

    def index_vault(self):
        """ボルト全体をインデックス化"""
        print("🔄 ボルトをインデックス中...")

        md_files = list(self.vault_path.rglob("*.md"))
        print(f"  発見: {len(md_files)} ファイル")

        indexed = 0
        for filepath in md_files:
            try:
                self._index_file(filepath)
                indexed += 1
            except Exception as e:
                print(f"  ⚠️ {filepath.name}: {e}")

        print(f"✓ インデックス完了: {indexed} ファイル")

    def _index_file(self, filepath: Path):
        """単一ファイルをインデックス化"""
        rel_path = filepath.relative_to(self.vault_path)

        with open(filepath, 'r', encoding='utf-8') as f:
            file_content = f.read()

        # フロントマターを手動で抽出
        metadata = {}
        content = file_content

        if file_content.startswith('---'):
            parts = file_content.split('---', 2)
            if len(parts) >= 3:
                try:
                    import yaml
                    metadata = yaml.safe_load(parts[1]) or {}
                    content = parts[2].strip()
                except:
                    pass

        # メタデータ抽出
        title = metadata.get('title', filepath.stem)
        doc_type = metadata.get('type', 'note')
        tags = metadata.get('tags', [])
        created_at = metadata.get('created_at', datetime.now().isoformat())

        # created_atがdate/datetimeオブジェクトの場合、文字列に変換
        if hasattr(created_at, 'isoformat'):
            created_at = created_at.isoformat()

        # JSON用の日付シリアライザー
        def json_serializer(obj):
            if hasattr(obj, 'isoformat'):
                return obj.isoformat()
            return str(obj)

        # データベースに記録
        cursor = self.conn.cursor()
        cursor.execute('''
        INSERT OR REPLACE INTO notes
        (filepath, title, content, type, created_at, modified_at, tags, metadata)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            str(rel_path),
            title,
            content,
            doc_type,
            created_at,
            datetime.now().isoformat(),
            json.dumps(tags if isinstance(tags, list) else [tags], default=json_serializer),
            json.dumps(metadata, default=json_serializer)
        ))

        note_id = cursor.lastrowid

        # タグを別テーブルに記録
        if isinstance(tags, list):
            for tag in tags:
                cursor.execute('INSERT INTO tags (note_id, tag) VALUES (?, ?)',
                             (note_id, tag))

        self.conn.commit()

    def search(self, query: str, tag_filter: str = None) -> List[Dict]:
        """
        検索を実行

        Args:
            query: 検索キーワード
            tag_filter: タグフィルター（オプション）

        Returns:
            検索結果リスト
        """
        cursor = self.conn.cursor()

        # 基本的な全文検索
        sql = '''
        SELECT DISTINCT n.* FROM notes n
        WHERE (n.title LIKE ? OR n.content LIKE ? OR n.tags LIKE ?)
        '''
        params = [f'%{query}%', f'%{query}%', f'%{query}%']

        # タグフィルターを適用
        if tag_filter:
            sql += ' AND n.tags LIKE ?'
            params.append(f'%{tag_filter}%')

        sql += ' ORDER BY n.modified_at DESC LIMIT 50'

        cursor.execute(sql, params)
        results = [dict(row) for row in cursor.fetchall()]

        return results

    def get_daily_summary(self, date: str = None) -> Dict:
        """
        指定日の日次要約を生成

        Args:
            date: ISO形式の日付（YYYY-MM-DD）。Noneなら今日

        Returns:
            日次要約辞書
        """
        if not date:
            date = datetime.now().strftime('%Y-%m-%d')

        cursor = self.conn.cursor()

        # その日のノート取得
        cursor.execute('''
        SELECT * FROM notes
        WHERE DATE(created_at) = ?
        ORDER BY created_at DESC
        ''', (date,))

        daily_notes = [dict(row) for row in cursor.fetchall()]

        return {
            'date': date,
            'total_notes': len(daily_notes),
            'notes': daily_notes,
            'summary': self._generate_summary(daily_notes)
        }

    def _generate_summary(self, notes: List[Dict]) -> str:
        """ノートリストから要約を生成"""
        if not notes:
            return "記録がありません"

        types = {}
        for note in notes:
            doc_type = note.get('type', 'note')
            types[doc_type] = types.get(doc_type, 0) + 1

        summary = f"{len(notes)}個のノート: " + ", ".join(
            f"{t}×{c}" for t, c in types.items()
        )
        return summary

    def close(self):
        """データベース接続を閉じる"""
        if self.conn:
            self.conn.close()


def main():
    """メイン処理"""
    vault_path = r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳"

    search_engine = ResearchOSSearch(vault_path)

    # ボルトをインデックス化
    search_engine.index_vault()

    # 例: 検索テスト
    print("\n🔍 検索テスト: '#synthesis'")
    results = search_engine.search('synthesis', tag_filter='#synthesis')
    for result in results[:5]:
        print(f"  ✓ {result['filepath']} ({result['type']})")

    # 今日の要約
    print("\n📋 本日の要約")
    summary = search_engine.get_daily_summary()
    print(f"  {summary['summary']}")

    search_engine.close()


if __name__ == '__main__':
    main()
