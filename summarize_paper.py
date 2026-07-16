import sys
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"

if VENV_SITE_PACKAGES.exists():
    sys.path.insert(0, str(VENV_SITE_PACKAGES))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from anthropic import Anthropic
from dotenv import load_dotenv
from pypdf import PdfReader

import llm_client


PAPER_DIR = WORKSPACE_DIR / "paper"
OUTPUT_DIR = WORKSPACE_DIR / "summarize_paper"
MODEL = "claude-sonnet-5"
MAX_CHARS = 120_000


def extract_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    pages = []
    for index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"\n\n--- page {index} ---\n{text}")
    return "".join(pages).strip()


def summarize_paper(client: Anthropic, pdf_path: Path, overwrite: bool = False) -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_file = OUTPUT_DIR / f"{pdf_path.stem}_summary.md"
    if output_file.exists() and not overwrite:
        print(f"既存の要約を使用: {output_file}")
        return output_file

    paper_text = extract_pdf_text(pdf_path)
    if not paper_text:
        raise ValueError(f"PDFからテキストを抽出できませんでした: {pdf_path}")

    if len(paper_text) > MAX_CHARS:
        paper_text = paper_text[:MAX_CHARS]

    prompt = f"""
あなたは研究支援AIです。以下のPDF抽出テキストを、日本語の研究ログとして読みやすいMarkdown形式で整理してください。

出力形式:
# 論文要約

- タイトル: ...
- 出典: ...
- PDFファイル名: ...
- 原文確認が必要: ...

## 1. 論文の概要
## 2. 著者が示した主張
## 3. 主な結果
## 4. 新規性
## 5. 実験や評価
## 6. 自分の研究に使えそうな点
## 7. 注意点・限界
## 8. キーワード

重要:
- 論文タイトルが分かる場合は「- タイトル: ...」として必ず書いてください。
- PDF抽出由来の文字化けや欠落がある場合は「原文確認が必要」と明記してください。
- 研究メモとして後から読み返しやすい粒度で要約してください。

PDFファイル名:
{pdf_path.name}

PDF抽出テキスト:
{paper_text}
"""

    summary = llm_client.create_response(client, MODEL, prompt)

    output_file.write_text(summary, encoding="utf-8-sig")
    print(f"要約を保存: {output_file}")
    return output_file


def summarize_all_papers(overwrite: bool = False) -> list[Path]:
    load_dotenv(WORKSPACE_DIR / ".env")

    # ユーザー方針(2026-07-09): 論文PDFのClaude要約は既定オフ（課金なし）。
    # 有効化は .env に PAPER_CLAUDE_SUMMARY=1 を設定する（明示的オプトイン）。
    import os
    if os.getenv("PAPER_CLAUDE_SUMMARY", "0").strip() != "1":
        raise RuntimeError(
            "論文PDFのClaude要約は既定で無効です（課金回避）。"
            "有効化するには .env に PAPER_CLAUDE_SUMMARY=1 を設定してください。"
        )
    client = llm_client.get_client()

    pdf_files = sorted(PAPER_DIR.glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError(f"{PAPER_DIR} フォルダにPDFがありません。")

    output_files = []
    for pdf_path in pdf_files:
        print(f"要約中: {pdf_path}")
        output_files.append(summarize_paper(client, pdf_path, overwrite=overwrite))
    return output_files


def main() -> None:
    overwrite = "--overwrite" in sys.argv
    output_files = summarize_all_papers(overwrite=overwrite)

    print("\n完了:")
    for output_file in output_files:
        print(output_file)


if __name__ == "__main__":
    main()
