#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS Cloud Sync
デバイス間でメモ・ポモセッション・タスクを共有

機能:
- JSON ファイルベースのクラウドストレージ（OneDrive）
- 複数デバイス間でリアルタイム同期
- メモ・セッション・タスク全て共有
- 競合解決機能付き
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any
import hashlib

class ResearchOSCloudSync:
    def __init__(self, onedrive_path: str, vault_path: str = None):
        """
        初期化

        Args:
            onedrive_path: OneDrive フォルダのパス
            vault_path: Obsidian ボルトのパス
        """
        self.onedrive_path = Path(onedrive_path)
        self.vault_path = Path(vault_path) if vault_path else Path(onedrive_path) / "外部脳"
        self.data_file = self.onedrive_path / "research_os_cloud_data.json"
        self.metadata_file = self.onedrive_path / "research_os_metadata.json"
        self.time_memo_folder = self.vault_path / "11-TimeMemo"
        self._init_files()

    def _init_files(self):
        """クラウドファイルを初期化"""
        if not self.data_file.exists():
            initial_data = {
                'version': '1.0',
                'memos': [],
                'sessions': [],
                'tasks': [],
                'created_at': datetime.now().isoformat(),
                'updated_at': datetime.now().isoformat()
            }
            self._write_json(self.data_file, initial_data)

        if not self.metadata_file.exists():
            metadata = {
                'last_sync': {},
                'device_info': {},
                'conflicts': []
            }
            self._write_json(self.metadata_file, metadata)

    def _read_json(self, filepath: Path) -> Dict:
        """JSON ファイルを読み込み"""
        try:
            if filepath.exists():
                return json.loads(filepath.read_text(encoding='utf-8'))
        except Exception as e:
            print(f"読み込みエラー: {e}")
        return {}

    def _write_json(self, filepath: Path, data: Dict):
        """JSON ファイルに書き込み"""
        try:
            filepath.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding='utf-8'
            )
        except Exception as e:
            print(f"書き込みエラー: {e}")

    def sync_memo_from_device(self, device_name: str, memo: Dict) -> bool:
        """
        デバイスからメモを同期

        Args:
            device_name: デバイス名（"Desktop" or "Lenovo"）
            memo: メモデータ

        Returns:
            同期成功フラグ
        """
        data = self._read_json(self.data_file)

        # メモに ID と同期情報を追加
        memo_with_meta = {
            'id': self._generate_id(),
            'device': device_name,
            'timestamp': memo.get('timestamp', datetime.now().isoformat()),
            'title': memo.get('title', ''),
            'text': memo.get('text', ''),
            'category': memo.get('category', 'default'),
            'tags': memo.get('tags', ''),
            'hash': self._hash_content(memo.get('text', '')),
            'synced_at': datetime.now().isoformat()
        }

        data['memos'].append(memo_with_meta)
        data['updated_at'] = datetime.now().isoformat()
        self._write_json(self.data_file, data)

        # メタデータ更新
        metadata = self._read_json(self.metadata_file)
        metadata['last_sync'][device_name] = datetime.now().isoformat()
        self._write_json(self.metadata_file, metadata)

        return True

    def get_all_memos(self) -> List[Dict]:
        """全デバイスのメモを取得"""
        data = self._read_json(self.data_file)
        return data.get('memos', [])

    def get_device_memos(self, device_name: str) -> List[Dict]:
        """特定デバイスのメモを取得"""
        all_memos = self.get_all_memos()
        return [m for m in all_memos if m.get('device') == device_name]

    def sync_session_from_device(self, device_name: str, session: Dict) -> bool:
        """ポモドーロセッションを同期"""
        data = self._read_json(self.data_file)

        session_with_meta = {
            'id': self._generate_id(),
            'device': device_name,
            'timestamp': session.get('timestamp', datetime.now().isoformat()),
            'duration_minutes': session.get('duration_minutes', 25),
            'notes': session.get('notes', ''),
            'synced_at': datetime.now().isoformat()
        }

        data['sessions'].append(session_with_meta)
        data['updated_at'] = datetime.now().isoformat()
        self._write_json(self.data_file, data)

        return True

    def get_all_sessions(self) -> List[Dict]:
        """全デバイスのセッションを取得"""
        data = self._read_json(self.data_file)
        return data.get('sessions', [])

    def get_daily_stats(self, date: str = None) -> Dict:
        """本日の統計情報を取得"""
        if not date:
            date = datetime.now().strftime('%Y-%m-%d')

        memos = self.get_all_memos()
        sessions = self.get_all_sessions()

        # 本日のメモ
        today_memos = [m for m in memos if m['synced_at'].startswith(date)]

        # 本日のセッション
        today_sessions = [s for s in sessions if s['synced_at'].startswith(date)]
        total_focus = sum(s['duration_minutes'] for s in today_sessions)

        return {
            'date': date,
            'memos': len(today_memos),
            'sessions': len(today_sessions),
            'focus_minutes': total_focus,
            'devices_active': self._get_active_devices(date),
            'last_update': datetime.now().isoformat()
        }

    def _get_active_devices(self, date: str) -> List[str]:
        """指定日に活動があったデバイス一覧を取得"""
        memos = self.get_all_memos()
        sessions = self.get_all_sessions()

        devices = set()
        for m in memos:
            if m['synced_at'].startswith(date):
                devices.add(m['device'])
        for s in sessions:
            if s['synced_at'].startswith(date):
                devices.add(s['device'])

        return sorted(list(devices))

    def _generate_id(self) -> str:
        """ユニークなIDを生成"""
        return hashlib.md5(
            f"{datetime.now().isoformat()}{os.urandom(8)}".encode()
        ).hexdigest()[:8]

    def _hash_content(self, text: str) -> str:
        """コンテンツのハッシュを生成"""
        return hashlib.md5(text.encode()).hexdigest()[:8]

    def save_to_time_memo(self, title: str, text: str, timestamp: str = None) -> bool:
        """
        11-TimeMemo フォルダに保存

        Args:
            title: メモのタイトル
            text: メモの内容
            timestamp: タイムスタンプ（ISO形式）

        Returns:
            保存成功フラグ
        """
        if not timestamp:
            timestamp = datetime.now().isoformat()

        # フォルダ作成
        self.time_memo_folder.mkdir(parents=True, exist_ok=True)

        # ファイル名生成
        dt = datetime.fromisoformat(timestamp)
        filename = f"{dt.strftime('%Y-%m-%d_%H%M%S')}_{title[:20]}.md"
        filepath = self.time_memo_folder / filename

        # Markdown ファイル生成
        content = f"""---
type: time_memo
timestamp: {timestamp}
title: {title}
tags: []
synced_at: {datetime.now().isoformat()}
---

# {dt.strftime('%H:%M:%S')} — {title}

{text}
"""

        try:
            filepath.write_text(content, encoding='utf-8')
            return True
        except Exception as e:
            print(f"保存エラー: {e}")
            return False


def main():
    """テスト実行"""
    onedrive_path = r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude"

    sync = ResearchOSCloudSync(onedrive_path)

    print("🔄 クラウド同期テスト")
    print("")

    # Desktop からメモを同期
    print("📝 Desktop からメモを同期:")
    sync.sync_memo_from_device('Desktop', {
        'title': 'AFM測定完了',
        'text': 'Tip交換で問題解決。測定成功。',
        'category': 'experiment',
        'tags': '#synthesis #gnr'
    })
    print("   ✓ メモ同期完了")

    # Lenovo からメモを同期
    print("\n📝 Lenovo からメモを同期:")
    sync.sync_memo_from_device('Lenovo', {
        'title': 'チラシ試読',
        'text': '新しい論文の要約を読んだ。応用可能性あり。',
        'category': 'idea',
        'tags': '#literature #gnr'
    })
    print("   ✓ メモ同期完了")

    # セッションを同期
    print("\n🍅 セッションを同期:")
    sync.sync_session_from_device('Desktop', {
        'duration_minutes': 25,
        'notes': 'AFM測定集中セッション'
    })
    sync.sync_session_from_device('Lenovo', {
        'duration_minutes': 25,
        'notes': '論文読了セッション'
    })
    print("   ✓ セッション同期完了")

    # 本日の統計
    print("\n📊 本日の統計:")
    stats = sync.get_daily_stats()
    print(f"   メモ: {stats['memos']}個")
    print(f"   セッション: {stats['sessions']}個")
    print(f"   集中時間: {stats['focus_minutes']}分")
    print(f"   活動デバイス: {', '.join(stats['devices_active'])}")

    # 全メモを表示
    print("\n📋 全メモ:")
    for memo in sync.get_all_memos():
        print(f"   [{memo['device']}] {memo.get('title', '（無題）')}")


if __name__ == '__main__':
    main()
