import argparse
import html
import re
import sys
import webbrowser
from pathlib import Path

from summarize_paper import OUTPUT_DIR, summarize_all_papers


DEFAULT_ONENOTE_URL = (
    "onenote:https://d.docs.live.net/A7339F0AE206A171/"
    "%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88/"
    "FarEasternTribe/paper_summarize.one"
    "#section-id={887E9B60-85CD-4C6C-B715-BB60C56BA9D7}&end"
)


def read_markdown_files() -> list[Path]:
    files = sorted(OUTPUT_DIR.glob("*_summary.md"))
    if not files:
        raise FileNotFoundError(f"{OUTPUT_DIR} に *_summary.md がありません。")
    return files


def extract_title(markdown: str, fallback: str) -> str:
    title_match = re.search(r"^\s*[-*]\s*タイトル\s*[:：]\s*(.+?)\s*$", markdown, re.MULTILINE)
    if title_match:
        return title_match.group(1).strip()

    for line in markdown.splitlines():
        line = line.strip()
        if line.startswith("# "):
            heading = line[2:].strip()
            if heading and heading != "論文要約":
                return heading
    return fallback


def markdown_to_html(markdown: str, page_title: str) -> str:
    body_parts: list[str] = []
    in_list = False

    def close_list() -> None:
        nonlocal in_list
        if in_list:
            body_parts.append("</ul>")
            in_list = False

    for raw_line in markdown.replace("\r\n", "\n").split("\n"):
        line = raw_line.rstrip()
        stripped = line.strip()

        if not stripped:
            close_list()
            continue

        if re.match(r"^#\s*論文要約\s*$", stripped):
            continue

        if re.match(r"^[-*]\s*タイトル\s*[:：]", stripped):
            continue

        heading = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading:
            close_list()
            level = min(len(heading.group(1)) + 1, 4)
            body_parts.append(f"<h{level}>{html.escape(heading.group(2))}</h{level}>")
            continue

        bullet = re.match(r"^[-*]\s+(.+)$", stripped)
        if bullet:
            if not in_list:
                body_parts.append("<ul>")
                in_list = True
            body_parts.append(f"<li>{html.escape(bullet.group(1))}</li>")
            continue

        numbered = re.match(r"^(\d+\).+)$", stripped)
        if numbered:
            if not in_list:
                body_parts.append("<ul>")
                in_list = True
            body_parts.append(f"<li>{html.escape(numbered.group(1))}</li>")
            continue

        close_list()
        body_parts.append(f"<p>{html.escape(stripped)}</p>")

    close_list()

    return f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>{html.escape(page_title)}</title>
  <style>
    body {{
      font-family: "Yu Gothic", "Meiryo", "Segoe UI", sans-serif;
      line-height: 1.75;
      max-width: 920px;
      margin: 40px auto;
      padding: 0 24px 80px;
      color: #1f2933;
      background: #ffffff;
    }}
    h1 {{
      font-size: 28px;
      line-height: 1.35;
      margin: 0 0 24px;
      padding-bottom: 12px;
      border-bottom: 2px solid #d0d7de;
    }}
    h2 {{
      font-size: 22px;
      margin: 32px 0 12px;
      padding-left: 10px;
      border-left: 5px solid #2563eb;
    }}
    h3, h4 {{
      font-size: 18px;
      margin: 24px 0 8px;
    }}
    p {{
      margin: 8px 0;
    }}
    ul {{
      margin: 8px 0 18px 1.3em;
      padding: 0;
    }}
    li {{
      margin: 6px 0;
    }}
    .note {{
      margin: 0 0 24px;
      padding: 12px 16px;
      background: #f6f8fa;
      border: 1px solid #d0d7de;
      border-radius: 8px;
      color: #57606a;
    }}
  </style>
</head>
<body>
  <h1>{html.escape(page_title)}</h1>
  <div class="note">このページをブラウザで開き、本文を選択してコピーすると、OneNoteへ比較的きれいに貼り付けられます。</div>
  {"".join(body_parts)}
</body>
</html>
"""


def write_html_file(md_path: Path) -> Path:
    markdown = md_path.read_text(encoding="utf-8-sig")
    title = extract_title(markdown, md_path.stem.replace("_summary", ""))
    html_file = md_path.with_suffix(".html")
    html_file.write_text(markdown_to_html(markdown, title), encoding="utf-8")
    return html_file


def build_html_files(md_files: list[Path]) -> list[Path]:
    html_files = []
    for md_path in md_files:
        html_file = write_html_file(md_path)
        html_files.append(html_file)
        print(f"HTMLを保存: {html_file}")
    return html_files


def open_outputs(html_files: list[Path], onenote_url: str | None, notion_url: str | None) -> None:
    for html_file in html_files:
        webbrowser.open(html_file.resolve().as_uri())
    if onenote_url:
        webbrowser.open(onenote_url)
    if notion_url:
        webbrowser.open(notion_url)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="paper内PDFを要約し、summarize_paperにMarkdownとOneNote貼り付け用HTMLを作成します。"
    )
    parser.add_argument("--skip-summarize", action="store_true", help="既存の *_summary.md からHTMLだけ作る")
    parser.add_argument("--summarize-only", action="store_true", help="PDF要約だけ行い、HTMLは作らない")
    parser.add_argument("--html-only", action="store_true", help="PDF要約後にHTMLを作るが、ブラウザでは開かない")
    parser.add_argument("--overwrite", action="store_true", help="既存の要約Markdownを上書きして再生成する")
    parser.add_argument("--no-onenote", action="store_true", help="OneNoteをブラウザで開かない")
    parser.add_argument("--onenote-url", default=DEFAULT_ONENOTE_URL, help="OneNoteのpaper_summarizeセクションURL")
    parser.add_argument("--notion-url", help="Notionの貼り付け先ページURL。指定するとHTMLと一緒に開く")
    args = parser.parse_args()

    try:
        if args.skip_summarize:
            md_files = read_markdown_files()
        else:
            md_files = summarize_all_papers(overwrite=args.overwrite)

        if args.summarize_only:
            return

        html_files = build_html_files(md_files)
        if not args.html_only:
            open_outputs(html_files, None if args.no_onenote else args.onenote_url, args.notion_url)

        print("\n完了:")
        print("1. 開いたHTMLページで本文を選択してコピー")
        print("2. OneNoteまたはNotionの新しいページに貼り付け")
        print("3. ページタイトルにはHTML上部の論文タイトルを使ってください")
    except Exception as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()
