---
title: Research OS - Personal Knowledge Management System
description: Obsidian-based full-text search and auto-classification engine for research
date_created: 2026-08-05
author: Claude (Anthropic)
status: v1.0 (Stable)
---

# Research OS 🧠

小島さん専用の研究支援 OS。Obsidian ボルト + Python 検索エンジン + 自動分類 + HTML ダッシュボード。

---

## 🎯 概要

**問題点**:
- 日常ログ、論文、実験ノートが散文的に管理されていた
- 357個のファイルから必要な情報を探すのに時間がかかる
- タイムライン入力の分類が手作業

**解決策**:
- Obsidian の構造を統一化（12フォルダ + タグスキーマ）
- SQLite 全文検索インデックス（2.6 MB）
- キーワードベースの自動分類エンジン
- リアルタイム司令画面（HTML ダッシュボード）

**成果**:
- 357 ファイルを秒速で検索可能
- タイムライン入力を自動分類
- 朝のスタートアップ画面で優先事項を確認

---

## 📁 ディレクトリ構成

```
research_tools/
├─ research_os_search.py           # 全文検索エンジン
├─ research_os_classifier.py       # タイムライン自動分類
├─ research_os_daily_summary.py   # 日次要約生成
├─ research_os_dashboard.html      # HTML司令画面
├─ research_os.db                  # SQLiteインデックス（自動生成）
├─ README.md                        # このファイル
└─ requirements.txt                 # Python依存パッケージ
```

---

## ⚙️ インストール

### 前提条件

- Python 3.8+
- Obsidian（外部脳ボルト）
- Windows 11 (PowerShell)

### セットアップ

#### 1. Python パッケージのインストール

```bash
cd research_tools
pip install -r requirements.txt
# または
pip install pyyaml
```

#### 2. Obsidian ボルトの準備

外部脳フォルダが以下の構造になっていることを確認:

```
外部脳/
├─ 00_Inbox/              # 未処理入力
├─ 01_Daily/              # 日常ログ
├─ 02_Projects/           # プロジェクト
├─ 03_Experiments/        # 実験記録
├─ 04_Papers/             # 論文
├─ 05_Ideas/              # アイデア・ChatGPT会話
├─ 06_Meetings/           # 会議記録
├─ 07_Health/             # 体調・睡眠ログ
├─ 08_Admin/              # 学務
├─ 09_People/             # プロフィール
├─ 10_Templates/          # テンプレート
└─ 99_Archive/            # アーカイブ
```

もし古い構造の場合、`migrate_to_new_structure.ps1` を実行:
```bash
.\migrate_to_new_structure.ps1
```

#### 3. 初回インデックス生成

```bash
python research_os_search.py
```

→ `research_os.db` が生成されます（初回は30秒程度）

---

## 🚀 使い方

### 1️⃣ ダッシュボードを開く

```bash
# ブラウザで開く
start research_os_dashboard.html

# または手動でファイルダブルクリック
```

**表示内容**:
- 今日の状態（睡眠時間、体調、集中可能時間）
- 優先事項（緊急・高・通常）
- 進行中のプロジェクト
- 本日のまとめ

### 2️⃣ タイムライン入力を自動分類

```bash
python research_os_classifier.py
```

**入力例**:
```
16:42 Scholl反応開始
16:55 FeCl3を追加
17:20 溶液が黒色化
17:45 STM roomへ移動
```

**自動処理**:
- キーワードマッチで分類（experiment / meeting / health / admin）
- タグ自動付与（#synthesis #gnr #in-progress等）
- Obsidian に自動保存（03_Experiments/ など）

### 3️⃣ 検索を実行

```bash
python -c "
from research_os_search import ResearchOSSearch
search = ResearchOSSearch(r'C:\Users\laput\...\外部脳')
results = search.search('Scholl反応', tag_filter='#synthesis')
for r in results:
    print(f\"{r['filepath']}: {r['type']}\")
search.close()
"
```

またはスクリプトとして:

```python
from research_os_search import ResearchOSSearch

vault_path = r"C:\Users\laput\...\外部脳"
search = ResearchOSSearch(vault_path)

# 検索
results = search.search("GNR合成")
for result in results[:5]:
    print(f"✓ {result['filepath']}")

search.close()
```

### 4️⃣ 日次要約を生成

```bash
python research_os_daily_summary.py
```

**出力例**:
```
📊 本日の要約を生成中...

✓ 生成完了: 2026-08-05
  総ノート数: 95
  進捗: 論文×46 | note×18 | 日常ログ×12 | ...

🏷️ タグ統計（上位5件）:
  論文: 46
  supporting information: 25
  experimental: 23
```

---

## 📋 タグスキーマ

**分野**:
- `#research` — 研究全般
- `#synthesis` — 有機合成実験
- `#characterization` — 構造解析（AFM, STM等）
- `#literature` — 論文調査

**状態**:
- `#todo` — 未着手
- `#in-progress` — 進行中
- `#completed` — 完了

**プロジェクト**:
- `#gnr` — グラフェンナノリボン
- `#phosphorus` — オレンジリン
- `#cage` — ケージ合成
- `#kaken` — 科研費

**優先度**:
- `#urgent` — 緊急（1-2日）
- `#high` — 高優先度（1週間）
- `#normal` — 通常
- `#low` — 低優先度

**健康・生活**:
- `#health` — 体調記録
- `#sleep` — 睡眠
- `#diet` — 食事
- `#stress` — ストレス

---

## 🔍 検索例

```python
# GNR合成の完了した実験を検索
results = search.search("", tag_filter="#synthesis #gnr #completed")

# 健康関連の緊急課題
results = search.search("", tag_filter="#health #urgent")

# 未完了の教育関連タスク
results = search.search("", tag_filter="#teaching #todo")

# キーワード + タグで検索
results = search.search("Scholl", tag_filter="#synthesis")
```

---

## 🛠️ トラブルシューティング

### Q. `research_os.db` が見つからない

**A.** 初回は手動で生成が必要です:
```bash
python research_os_search.py
```

### Q. ダッシュボードを開くと真っ白

**A.** ブラウザをリロード (Ctrl+R) または JavaScript コンソール (F12) でエラーを確認。

### Q. 分類がうまくいかない

**A.** キーワードを `research_os_classifier.py` の `self.rules` に追加:
```python
self.rules = {
    'experiment': {
        'keywords': [..., '新しいキーワード'],  # ここに追加
        ...
    }
}
```

### Q. Obsidian で新しいタグが反映されない

**A.** Obsidian を再起動するか、以下を実行してインデックスを再生成:
```bash
del research_os.db
python research_os_search.py
```

---

## 📊 技術仕様

### アーキテクチャ

```
Focus Desk PWA (タイムライン入力)
    ↓
[HH:MM テキスト入力]
    ↓
research_os_classifier.py
    ↓
[自動分類・フロントマター付与]
    ↓
Obsidian フォルダに自動保存
    ↓
research_os_search.py
    ↓
[SQLite インデックス化]
    ↓
research_os_dashboard.html
    ↓
[リアルタイム表示]
```

### データベース構成

**SQLite テーブル**:

1. **notes**: ファイル情報
   ```sql
   id, filepath, title, content, type, created_at, 
   modified_at, tags, metadata, indexed_at
   ```

2. **search_index**: 単語インデックス
   ```sql
   id, note_id, word, frequency
   ```

3. **tags**: タグ対応
   ```sql
   id, note_id, tag
   ```

### パフォーマンス

- インデックス生成: ~2秒 (357ファイル)
- 検索クエリ: <100ms
- DB容量: 2.64 MB
- メモリ使用: ~50 MB

---

## 📈 拡張性

### 次のステップ（9月以降）

- [ ] ベクトル検索（意味検索）実装
- [ ] Obsidian Sync との統合
- [ ] スマホダッシュボード
- [ ] エージェント分業（AI役割分離）
- [ ] 論文 PDF 自動抽出

### カスタマイズ

**分類ルールの追加**:
```python
self.rules['custom_category'] = {
    'keywords': ['キーワード1', 'キーワード2'],
    'destination': '10_Templates',
    'template': None
}
```

**タグの追加**:
`10_Templates/TAG_SCHEMA.md` を編集

---

## 📝 ファイル説明

### research_os_search.py

**用途**: 全文検索エンジン

**クラス**: `ResearchOSSearch`

**主なメソッド**:
- `index_vault()` — ボルト全体をインデックス化
- `search(query, tag_filter)` — 検索実行
- `get_daily_summary(date)` — 日次要約生成

**実行**:
```bash
python research_os_search.py
```

---

### research_os_classifier.py

**用途**: タイムライン自動分類

**クラス**: `TimelineClassifier`

**分類カテゴリ**:
- experiment（合成・計測）
- meeting（ゼミ・会議）
- health（体調・睡眠）
- admin（学務・事務）
- default（その他）

**実行**:
```bash
python research_os_classifier.py
```

---

### research_os_daily_summary.py

**用途**: 日次要約生成

**クラス**: `DailySummaryGenerator`

**出力**: HTML 形式のサマリー

**実行**:
```bash
python research_os_daily_summary.py
```

---

### research_os_dashboard.html

**用途**: リアルタイム司令画面

**表示**:
- 今日の状態（睡眠・体調・集中度）
- 優先事項（優先度別）
- 進行中プロジェクト
- 本日のまとめ

**ブラウザで開く**:
```bash
start research_os_dashboard.html
```

---

## 🔐 セキュリティ・プライバシー

- ✅ データはすべてローカル管理（クラウドに送信なし）
- ✅ OneDrive で自動バックアップ
- ✅ git で変更履歴を記録
- ⚠️ `research_os.db` は git 対象外（`.gitignore` に記載）

---

## 📄 ライセンス

このプロジェクトは小島さん専用です。

将来的に他の研究者と共有する場合は MIT License 等を適用してください。

---

## 📞 サポート・改善提案

Claude Code セッションで随時改善・拡張可能です。

**よくある改善要望**:
- 新しい分類カテゴリの追加
- ダッシュボードのカスタマイズ
- パフォーマンス最適化
- Obsidian プラグイン化

---

## 🎓 参考資料

- [Obsidian 公式](https://obsidian.md/)
- [SQLite 全文検索](https://www.sqlite.org/fts5.html)
- [Python frontmatter](https://python-frontmatter.readthedocs.io/)

---

**最終更新**: 2026-08-05 17:30  
**バージョン**: 1.0 (Stable)  
**開発者**: Claude (Anthropic)
