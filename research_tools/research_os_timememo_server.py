#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Research OS TimeMemo HTTP Server
ダッシュボードから HTTP POST でメモを受け取り、11-TimeMemo に直接保存
"""

from flask import Flask, request, jsonify
from pathlib import Path
from datetime import datetime
import json
import sys

app = Flask(__name__)

# パス設定
ONEDRIVE_PATH = Path(r"C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude")
VAULT_PATH = ONEDRIVE_PATH / "外部脳"
TIMEMEMO_FOLDER = VAULT_PATH / "11-TimeMemo"

# フォルダ作成
TIMEMEMO_FOLDER.mkdir(parents=True, exist_ok=True)

@app.route('/save-timememo', methods=['POST'])
def save_timememo():
    """11-TimeMemo にメモを保存"""
    try:
        data = request.get_json()

        if not data or 'text' not in data:
            return jsonify({'error': 'メモテキストが必要です'}), 400

        text = data['text'].strip()
        if not text:
            return jsonify({'error': 'メモが空です'}), 400

        # ファイル名生成
        timestamp = datetime.now()
        title = data.get('title', text[:20])
        filename = f"{timestamp.strftime('%Y-%m-%d_%H%M%S')}_{title[:15]}.md"

        # frontmatter 付き Markdown 生成
        category = data.get('category', 'default')
        tags = data.get('tags', '')

        md_content = f"""---
type: time_memo
timestamp: {timestamp.isoformat()}
title: {title}
category: {category}
tags: [{', '.join(f"'{t}'" for t in tags.split() if t)}]
synced_at: {datetime.now().isoformat()}
---

# {timestamp.strftime('%H:%M:%S')} — {title}

{text}

## メタデータ

- 時刻: {timestamp.strftime('%H:%M:%S')}
- 日付: {timestamp.strftime('%Y-%m-%d')}
- カテゴリ: {category}
- 同期: ✓ Dashboard → Python Server → Obsidian
"""

        # ファイル保存
        filepath = TIMEMEMO_FOLDER / filename
        filepath.write_text(md_content, encoding='utf-8')

        return jsonify({
            'success': True,
            'message': f'保存完了: {filename}',
            'filepath': str(filepath)
        }), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    """ヘルスチェック"""
    return jsonify({'status': 'running', 'timememo_folder': str(TIMEMEMO_FOLDER)}), 200

if __name__ == '__main__':
    print(f"🚀 Research OS TimeMemo Server 起動")
    print(f"📁 保存先: {TIMEMEMO_FOLDER}")
    print(f"🌐 http://localhost:5000")
    print(f"ℹ  ダッシュボードから POST /save-timememo に送信")
    print()

    app.run(host='127.0.0.1', port=5000, debug=False)
