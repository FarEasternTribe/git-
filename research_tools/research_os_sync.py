#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS Sync Engine
ダッシュボード ↔ Obsidian 自動同期

機能:
- メモ自動保存
- ポモドーロセッション記録
- タスク状態同期
- 複数デバイス同期対応
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import sqlite3

class ResearchOSSync:
    def __init__(self, vault_path: str, db_path: str = "research_os_sync.db"):
        self.vault_path = Path(vault_path)
        self.db_path = Path(db_path)
        self._init_database()

    def _init_database(self):
        """同期用SQLiteデータベースを初期化"""
        self.conn = sqlite3.connect(str(self.db_path))
        cursor = self.conn.cursor()

        # メモテーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS memos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            title TEXT,
            content TEXT NOT NULL,
            category TEXT,
            tags TEXT,
            status TEXT DEFAULT 'pending',
            obsidian_file TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            synced_at TEXT
        )
        ''')

        # ポモドーロセッションテーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS pomodoro_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            duration_minutes INTEGER,
            completed INTEGER DEFAULT 1,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            synced_at TEXT
        )
        ''')

        # タスクテーブル
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            text TEXT NOT NULL,
            priority TEXT,
            completed INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            synced_at TEXT
        )
        ''')

        self.conn.commit()

    def save_memo(self, title: str, content: str, category: str, tags: str) -> int:
        """メモを保存"""
        cursor = self.conn.cursor()
        timestamp = datetime.now().isoformat()

        cursor.execute('''
        INSERT INTO memos (timestamp, title, content, category, tags)
        VALUES (?, ?, ?, ?, ?)
        ''', (timestamp, title, content, category, tags))

        self.conn.commit()
        memo_id = cursor.lastrowid

        # Obsidianに自動保存
        self._sync_memo_to_obsidian(memo_id)

        return memo_id

    def _sync_memo_to_obsidian(self, memo_id: int):
        """メモをObsidianに同期"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM memos WHERE id = ?', (memo_id,))
        memo = cursor.fetchone()

        if not memo:
            return

        _, timestamp, title, content, category, tags, _, obsidian_file, _, _ = memo

        # ファイル名生成
        now = datetime.now()
        filename = f"{now.strftime('%Y-%m-%d_%H%M%S')}_memo.md"

        # カテゴリーに応じたフォルダ決定
        folder_map = {
            'experiment': '03_Experiments',
            'idea': '05_Ideas',
            'meeting': '06_Meetings',
            'health': '07_Health',
            'default': '01_Daily'
        }
        folder = folder_map.get(category, '01_Daily')

        filepath = self.vault_path / folder / filename

        # フロントマター付きで保存
        frontmatter = f"""---
type: memo
timestamp: {timestamp}
title: {title if title else 'Untitled'}
category: {category}
tags: {tags}
synced_at: {datetime.now().isoformat()}
---

# {title if title else now.strftime('%Y-%m-%d %H:%M')}

{content}
"""

        filepath.write_text(frontmatter, encoding='utf-8')

        # DBを更新
        cursor.execute('''
        UPDATE memos SET obsidian_file = ?, status = 'synced', synced_at = ?
        WHERE id = ?
        ''', (str(filepath.relative_to(self.vault_path)), datetime.now().isoformat(), memo_id))

        self.conn.commit()

    def save_pomodoro_session(self, duration_minutes: int, notes: str = ""):
        """ポモドーロセッションを記録"""
        cursor = self.conn.cursor()
        date = datetime.now().strftime('%Y-%m-%d')

        cursor.execute('''
        INSERT INTO pomodoro_sessions (date, duration_minutes, notes)
        VALUES (?, ?, ?)
        ''', (date, duration_minutes, notes))

        self.conn.commit()
        session_id = cursor.lastrowid

        # Obsidianに同期
        self._sync_pomodoro_to_obsidian(session_id)

        return session_id

    def _sync_pomodoro_to_obsidian(self, session_id: int):
        """ポモドーロセッションをObsidianに同期"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM pomodoro_sessions WHERE id = ?', (session_id,))
        session = cursor.fetchone()

        if not session:
            return

        _, date, duration, completed, notes, _, _ = session

        # 日別の07_Health に記録
        date_obj = datetime.strptime(date, '%Y-%m-%d')
        filepath = self.vault_path / '07_Health' / f"{date}_pomodoro.md"

        content = f"""---
type: pomodoro_log
date: {date}
duration_minutes: {duration}
completed: {completed}
synced_at: {datetime.now().isoformat()}
---

# ポモドーロセッション - {date}

- 時間: {duration}分
- 完了: {'✓ はい' if completed else '✗ いいえ'}
{f'- メモ: {notes}' if notes else ''}
"""

        # 日付ファイルに追記する方式もあり
        daily_file = self.vault_path / '07_Health' / f"{date}.md"
        if daily_file.exists():
            existing = daily_file.read_text(encoding='utf-8')
            daily_file.write_text(
                existing + f"\n\n## ポモドーロ\n- {duration}分間集中 ({datetime.now().strftime('%H:%M')})\n",
                encoding='utf-8'
            )
        else:
            daily_file.write_text(content, encoding='utf-8')

        cursor.execute('''
        UPDATE pomodoro_sessions SET synced_at = ?
        WHERE id = ?
        ''', (datetime.now().isoformat(), session_id))

        self.conn.commit()

    def get_daily_summary(self, date: str = None) -> Dict:
        """本日のサマリーを取得"""
        if not date:
            date = datetime.now().strftime('%Y-%m-%d')

        cursor = self.conn.cursor()

        # メモ数
        cursor.execute(
            'SELECT COUNT(*) FROM memos WHERE DATE(timestamp) = ?', (date,)
        )
        memo_count = cursor.fetchone()[0]

        # ポモドーロセッション
        cursor.execute(
            'SELECT SUM(duration_minutes) FROM pomodoro_sessions WHERE date = ? AND completed = 1',
            (date,)
        )
        total_focus = cursor.fetchone()[0] or 0

        # タスク完了数
        cursor.execute(
            'SELECT COUNT(*) FROM tasks WHERE completed = 1 AND DATE(created_at) = ?',
            (date,)
        )
        task_completed = cursor.fetchone()[0]

        return {
            'date': date,
            'memos': memo_count,
            'focus_minutes': total_focus,
            'tasks_completed': task_completed,
            'sync_status': 'synced'
        }

    def close(self):
        """データベース接続を閉じる"""
        if self.conn:
            self.conn.close()


def main():
    """テスト実行"""
    vault_path = r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\外部脳"

    sync = ResearchOSSync(vault_path)

    # テスト: メモ保存
    print("🔄 同期テスト中...")
    memo_id = sync.save_memo(
        title="AFM測定完了",
        content="Tip交換により問題解決。測定成功。",
        category="experiment",
        tags="#synthesis #gnr #completed"
    )
    print(f"✓ メモ保存 ID: {memo_id}")

    # テスト: ポモドーロセッション保存
    session_id = sync.save_pomodoro_session(
        duration_minutes=25,
        notes="AFM測定集中セッション"
    )
    print(f"✓ ポモドーロセッション保存 ID: {session_id}")

    # 本日のサマリー
    summary = sync.get_daily_summary()
    print(f"\n📊 本日のサマリー:")
    print(f"  メモ: {summary['memos']}個")
    print(f"  集中時間: {summary['focus_minutes']}分")
    print(f"  タスク完了: {summary['tasks_completed']}個")

    sync.close()


if __name__ == '__main__':
    main()
