import json
import os
import re
import sys
import time
import argparse
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"

if VENV_SITE_PACKAGES.exists():
    sys.path.insert(0, str(VENV_SITE_PACKAGES))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from anthropic import Anthropic

from agent_config import load_agent_env

import llm_client


JOURNALS = {
    "Nature": "0028-0836",
    "Science": "0036-8075",
    "JACS": "0002-7863",
    "Nature Chemistry": "1755-4330",
    "Nature Communications": "2041-1723",
    "Nature Catalysis": "2520-1158",
    "Nature Materials": "1476-1122",
    "Nature Nanotechnology": "1748-3387",
    "ACS Nano": "1936-0851",
    "Nano Letters": "1530-6984",
    "Advanced Materials": "0935-9648",
    "J. Phys. Chem. Lett.": "1948-7185",
}

KEYWORDS = [
    # core
    "graphene nanoribbon",
    "GNR",
    "nanographene",
    "on-surface synthesis",
    "surface-assisted synthesis",
    "STM",
    "scanning tunneling microscopy",
    "Au(111)",

    # CO2RR / electrocatalysis
    "CO2RR",
    "CO2 reduction",
    "carbon dioxide reduction",
    "electrocatalysis",
    "oxygen evolution reaction",
    "OER",
    "hydrogen evolution reaction",
    "HER",

    # chirality / spin
    "chirality",
    "CISS",
    "spin polarization",
    "spin selectivity",
    "spintronics",

    # materials
    "2D materials",
    "graphene",
    "MXene",
    "COF",
    "covalent organic framework",
    "MOF",
    "metal organic framework",
    "single atom catalyst",

    # Si / etching
    "silicon etching",
    "chemical etching",
    "metal assisted chemical etching",
    "MACE",
]


def parse_args():
    parser = argparse.ArgumentParser(description="Daily paper search and Claude relevance scoring.")
    parser.add_argument("--days", type=int, default=3, help="何日前まで検索するか。例: --days 7")
    parser.add_argument("--no-gpt", action="store_true", help="Claude関連度判定をスキップする")
    parser.add_argument(
        "--use-claude",
        action="store_true",
        help="Claude関連度判定を有効化する（既定は課金回避のため無効）。",
    )
    parser.add_argument("--max-gpt", type=int, default=30, help="Claudeで判定する最大論文数")
    parser.add_argument(
        "--score-batch",
        type=int,
        default=SCORING_BATCH_SIZE,
        help=f"Claude判定を何件ずつまとめて送るか（トークン節約）。省略時は {SCORING_BATCH_SIZE}",
    )
    return parser.parse_args()


def clean_html(text: str) -> str:
    text = text or ""
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def crossref_get(params: dict) -> dict:
    query = urllib.parse.urlencode(params)
    url = f"https://api.crossref.org/works?{query}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Claude-Agent daily_paper_search.py (mailto:example@example.com)",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))


def get_published_date(item: dict) -> str:
    for key in ["published-online", "published-print", "published"]:
        if key in item:
            parts = item[key]["date-parts"][0]
            return "-".join(map(str, parts))
    return ""


# 関連度スコアリングは軽い分類タスクなので Haiku で十分（ユーザー承認 2026-07-15、コスト減）。
# PAPER_SCORING_MODEL 環境変数で上書き可能（例: claude-sonnet-5 に戻す）。
SCORING_MODEL = os.getenv("PAPER_SCORING_MODEL", "claude-haiku-4-5-20251001")

# 固定の前置き（先生の関心）。バッチ内の全論文で共有し、論文ごとに再送しない＝トークン節約。
SCORING_PREAMBLE = """あなたは小島先生専属の論文秘書AIです。

小島先生の関心:
- グラフェンナノリボン GNR
- ナノグラフェン
- STM / 表面科学 / Au(111)
- on-surface synthesis
- CO2RR / CO2 reduction / electrocatalysis
- OER / HER
- キラリティ / CISS / スピン偏極
- 2D materials / MXene / MOF / COF
- Si chemical etching / MACE
- AI for Science
"""

# 長いアブストラクトは入力トークン節約のため切り詰める（関連度判定には十分）。
MAX_ABSTRACT_CHARS = 1200
# 1回のAPI呼び出しでまとめて判定する論文数（前置きの再送回数を減らす）。PAPER_SCORING_BATCHで上書き可。
SCORING_BATCH_SIZE = max(1, int(os.getenv("PAPER_SCORING_BATCH", "10")))


def _paper_block(row: dict, index: int | None = None) -> str:
    abstract = (row.get("Abstract", "") or "")[:MAX_ABSTRACT_CHARS]
    head = f"[{index}]\n" if index is not None else ""
    return (
        f"{head}Journal: {row.get('Journal', '')}\n"
        f"Matched keyword: {row.get('Keyword matched', '')}\n"
        f"Title: {row.get('Title', '')}\n"
        f"Abstract: {abstract}"
    )


def _to_columns(data: dict) -> dict:
    return {
        "Relevance score": data.get("score", ""),
        "Claude reason": data.get("reason", ""),
        "Research connection": data.get("research_connection", ""),
        "Recommended action": data.get("action", ""),
        "Claude keywords": ", ".join(data.get("keywords", []) or []),
    }


def _strip_json_fences(text: str) -> str:
    return text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()


def _error_columns(exc: object) -> dict:
    cols = {k: "" for k in _to_columns({})}
    cols["Claude reason"] = f"Claude判定エラー: {exc}"
    return cols


def score_with_claude(client: Anthropic, row: dict) -> dict:
    """1論文だけを判定する（バッチ失敗時のフォールバック）。"""
    prompt = (
        SCORING_PREAMBLE
        + "\n以下の論文が小島先生の研究にどれくらい関連するか判定してください。\n"
        "必ず下記形式のJSONだけで返してください（説明文不要）。\n"
        '{"score": 0から100の整数, "reason": "理由を日本語で短く", '
        '"research_connection": "研究との接点", '
        '"action": "読むべき / 要旨だけ確認 / 保留 / 無視", "keywords": ["関連キーワード", "..."]}\n\n'
        + _paper_block(row)
    )
    try:
        data = json.loads(_strip_json_fences(llm_client.create_response(client, SCORING_MODEL, prompt)))
        return _to_columns(data)
    except Exception as exc:
        return _error_columns(exc)


def _score_batch(client: Anthropic, batch: list[dict]) -> list[dict]:
    if len(batch) == 1:
        return [score_with_claude(client, batch[0])]
    papers = "\n\n".join(_paper_block(row, i) for i, row in enumerate(batch))
    prompt = (
        SCORING_PREAMBLE
        + f"\n以下の{len(batch)}件の論文それぞれについて、小島先生の研究との関連度を判定してください。\n"
        "必ずJSON配列だけで返してください（説明文不要）。各要素は次の形式:\n"
        '{"index": 論文番号[n], "score": 0から100の整数, "reason": "理由を日本語で短く", '
        '"research_connection": "研究との接点", '
        '"action": "読むべき / 要旨だけ確認 / 保留 / 無視", "keywords": ["関連キーワード", "..."]}\n\n'
        "論文:\n" + papers
    )
    try:
        data = json.loads(_strip_json_fences(llm_client.create_response(client, SCORING_MODEL, prompt)))
        if not isinstance(data, list):
            raise ValueError("expected a JSON array")
        by_index: dict[int, dict] = {}
        for n, item in enumerate(data):
            if not isinstance(item, dict):
                continue
            try:
                by_index[int(item.get("index", n))] = item
            except (TypeError, ValueError):
                by_index[n] = item
        # 揃った分はそのまま採用。欠けた論文だけ個別に再判定する。
        return [
            _to_columns(by_index[i]) if i in by_index else score_with_claude(client, row)
            for i, row in enumerate(batch)
        ]
    except Exception:
        # バッチ全体が失敗したら1件ずつにフォールバック（従来と同じ結果に劣化するだけ）。
        return [score_with_claude(client, row) for row in batch]


def score_papers_with_claude(
    client: Anthropic, rows: list[dict], batch_size: int = SCORING_BATCH_SIZE
) -> list[dict]:
    """複数論文をまとめて判定してトークンを節約する。前置きはバッチ内で1回だけ送られる。"""
    bs = max(1, batch_size)
    total = len(rows)
    results: list[dict] = []
    for start in range(0, total, bs):
        batch = rows[start:start + bs]
        print(f"Claude scoring {start + 1}-{start + len(batch)}/{total} ...")
        results.extend(_score_batch(client, batch))
    return results


def main():
    args = parse_args()
    load_agent_env()

    today = date.today()
    from_date = today - timedelta(days=args.days)
    to_date = today

    output_dir = WORKSPACE_DIR / "papers"
    output_dir.mkdir(exist_ok=True)

    output_file = output_dir / f"{today}_paper_list.xlsx"

    rows = []
    seen_dois = set()

    for journal_name, issn in JOURNALS.items():
        for keyword in KEYWORDS:
            print(f"Searching: {journal_name} / {keyword}")

            params = {
                "filter": f"issn:{issn},from-pub-date:{from_date},until-pub-date:{to_date},type:journal-article",
                "query.bibliographic": keyword,
                "rows": 20,
                "select": "DOI,title,container-title,published-print,published-online,published,URL,abstract,author",
            }

            try:
                data = crossref_get(params)
                items = data["message"]["items"]

                for item in items:
                    doi = item.get("DOI", "")
                    if not doi or doi in seen_dois:
                        continue
                    seen_dois.add(doi)

                    title = item.get("title", [""])[0]
                    journal = item.get("container-title", [journal_name])[0]
                    url = item.get("URL", "")
                    abstract = clean_html(item.get("abstract", ""))
                    published = get_published_date(item)

                    authors = item.get("author", [])
                    author_text = ", ".join(
                        f"{a.get('given', '')} {a.get('family', '')}".strip()
                        for a in authors[:5]
                    )

                    rows.append(
                        {
                            "Date found": str(today),
                            "Journal": journal,
                            "Published": published,
                            "Keyword matched": keyword,
                            "Title": title,
                            "Authors": author_text,
                            "DOI": doi,
                            "URL": url,
                            "Abstract": abstract,
                        }
                    )

            except Exception as exc:
                print(f"Error: {journal_name} / {keyword}: {exc}")

            time.sleep(0.2)

    df = pd.DataFrame(rows)

    if df.empty:
        print("該当論文は見つかりませんでした。")
        return

    # ユーザー方針(2026-07-09): 論文のClaude関連度判定は既定オフ（課金なし）。
    # 有効化は .env に PAPER_CLAUDE_SCORING=1、または実行時 --use-claude。
    claude_scoring = (not args.no_gpt) and (
        args.use_claude or os.getenv("PAPER_CLAUDE_SCORING", "0").strip() == "1"
    )
    if (not args.no_gpt) and not claude_scoring:
        print(
            "[skip] Claude関連度判定をスキップしました（課金なし）。"
            "有効化するには .env に PAPER_CLAUDE_SCORING=1 を設定するか --use-claude を付けてください。"
        )
    if claude_scoring:
        client = llm_client.get_client()

        rows_to_score = [row.to_dict() for _, row in df.head(args.max_gpt).iterrows()]
        scored_rows = score_papers_with_claude(client, rows_to_score, batch_size=args.score_batch)
        score_df = pd.DataFrame(scored_rows)

        for col in score_df.columns:
            df.loc[df.index[: len(score_df)], col] = score_df[col].values

    if "Relevance score" in df.columns:
        df["Relevance score"] = pd.to_numeric(df["Relevance score"], errors="coerce")
        df = df.sort_values(["Relevance score", "Journal", "Published"], ascending=[False, True, False])
    else:
        df = df.sort_values(["Journal", "Published", "Title"])

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="papers")

        workbook = writer.book
        worksheet = writer.sheets["papers"]

        worksheet.freeze_panes = "A2"

        widths = {
            "A": 14,
            "B": 24,
            "C": 14,
            "D": 24,
            "E": 60,
            "F": 34,
            "G": 24,
            "H": 45,
            "I": 80,
            "J": 16,
            "K": 50,
            "L": 50,
            "M": 18,
            "N": 36,
        }

        for col, width in widths.items():
            worksheet.column_dimensions[col].width = width

    print(f"\n保存しました: {output_file}")
    print(f"件数: {len(df)}")


if __name__ == "__main__":
    main()