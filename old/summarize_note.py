import argparse
import sys
from datetime import date
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)

for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from openai import OpenAI


DEFAULT_INPUT = "rawtext"
DEFAULT_OUTPUT_DIR = "日誌"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="一日のテキストメモを日付付きの日誌Markdownとして整理します。"
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=DEFAULT_INPUT,
        help=f"読み込むテキストメモ、またはテキストメモ入りフォルダ。省略時は {DEFAULT_INPUT}",
    )
    parser.add_argument(
        "--date",
        default=date.today().isoformat(),
        help="日誌の日付。例: 2026-06-30。省略時は今日",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"出力先フォルダ。省略時は {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="同名の日誌ファイルがある場合に上書きする",
    )
    return parser.parse_args()


def build_prompt(note: str, journal_date: str) -> str:
    return f"""
あなたは、研究者の日誌・研究メモ・生活ログを整理する専属AI秘書です。
以下の一日のテキストメモを、後から読み返しやすい日本語の日誌Markdownに整理してください。

重要な方針:
- 原文に誤字・脱字・音声入力由来の変換ミスがある場合は、自然な日本語に修正してください。
- ただし、意味が不明な部分は勝手に補わず、「意味不明瞭」「確認したいこと」として分けてください。
- 事実、推測、ChatGPTによる分析を明確に分けてください。
- メモに書かれていないことを事実として断定しないでください。
- 研究、生活、体調、食事、対人、気持ち、翌日の確認事項を読みやすく整理してください。
- 箇条書き中心でよいですが、日誌として自然に読めるようにしてください。
- 最後の「ChatGPTによる詳細分析」は長くても構いません。できるだけ詳しく、深く分析してください。
- 分析では、研究の進め方、心理状態、生活リズム、対人ストレス、翌日の行動設計まで含めてください。
- 必要なら「仮説」「リスク」「次の一手」に分けてください。

出力形式:

# {journal_date} 生活・研究ログ整理

## 1. 誤字修正後の要約
原文の意味を保ちながら、誤字・脱字・表記ゆれを修正して、短く自然な文章でまとめてください。

## 2. 研究
- 今日行った研究・実験・解析
- 得られた結果
- 未完了のこと
- 次に確認すべきこと

## 3. 身体の状態
- 睡眠
- 疲労
- 体調
- 気分への影響

## 4. 食事
- 食べたもの
- 食欲
- 体調との関係

## 5. 生活・移動
- 移動
- 作業環境
- 時間の使い方

## 6. 対人・気持ちの整理
- 対人関係で起きたこと
- 感情の動き
- ストレス要因
- 事実と解釈の分離

## 7. 自己分析
- 今日の自分の行動パターン
- 集中できた点
- 乱れた点
- 繰り返し出ている傾向

## 8. 明日以降のTODO
- 研究
- 生活
- 連絡・事務
- 体調管理

## 9. 確認したいこと
意味が曖昧な記述、追加で確認すべき点を列挙してください。

## 10. ChatGPTによる詳細分析
以下を含めて、できるだけ詳しく分析してください。

### 10.1 今日の全体像
### 10.2 研究面の分析
### 10.3 生活・体調面の分析
### 10.4 心理面の分析
### 10.5 対人関係・ストレス構造の分析
### 10.6 明日に向けた行動戦略
### 10.7 注意すべきリスク
### 10.8 良かった点
### 10.9 次の一手

一日のテキストメモ:
{note}
"""


def read_note_text(input_path: Path) -> str:
    if input_path.is_file():
        return input_path.read_text(encoding="utf-8-sig")

    if not input_path.is_dir():
        raise FileNotFoundError(f"入力ファイルまたはフォルダが見つかりません: {input_path}")

    text_files = sorted(input_path.glob("*.txt"))
    if not text_files:
        raise FileNotFoundError(f"フォルダ内に .txt ファイルがありません: {input_path}")

    chunks = []
    for text_file in text_files:
        content = text_file.read_text(encoding="utf-8-sig").strip()
        if not content:
            continue
        chunks.append(f"--- {text_file.name} ---\n{content}")

    if not chunks:
        raise ValueError(f"フォルダ内の .txt ファイルがすべて空です: {input_path}")

    return "\n\n".join(chunks)


def unique_output_path(output_dir: Path, journal_date: str, overwrite: bool) -> Path:
    output_file = output_dir / f"{journal_date}_日誌.md"
    if overwrite or not output_file.exists():
        return output_file

    index = 2
    while True:
        candidate = output_dir / f"{journal_date}_日誌_{index}.md"
        if not candidate.exists():
            return candidate
        index += 1


def main() -> None:
    args = parse_args()
    load_dotenv(WORKSPACE_DIR / ".env")

    input_path = Path(args.input)
    if not input_path.is_absolute():
        input_path = WORKSPACE_DIR / input_path

    note = read_note_text(input_path)
    if not note.strip():
        raise ValueError(f"入力メモが空です: {input_path}")

    client = OpenAI()
    response = client.responses.create(
        model="gpt-5",
        input=build_prompt(note, args.date),
    )

    summary = response.output_text.strip() + "\n"

    output_dir = Path(args.output_dir)
    if not output_dir.is_absolute():
        output_dir = WORKSPACE_DIR / output_dir
    output_dir.mkdir(exist_ok=True)

    output_file = unique_output_path(output_dir, args.date, args.overwrite)
    output_file.write_text(summary, encoding="utf-8-sig")

    print(summary)
    print(f"\n保存しました: {output_file}")


if __name__ == "__main__":
    main()