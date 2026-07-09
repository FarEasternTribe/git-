import json
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
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)

for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI


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
    parser = argparse.ArgumentParser(description="Daily paper search and GPT relevance scoring.")
    parser.add_argument("--days", type=int, default=3, help="何日前まで検索するか。例: --days 7")
    parser.add_argument("--no-gpt", action="store_true", help="GPT関連度判定をスキップする")
    parser.add_argument("--max-gpt", type=int, default=30, help="GPT判定する最大論文数")
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
            "User-Agent": "OpenAI-Agent daily_paper_search.py (mailto:example@example.com)",
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


def score_with_gpt(client: OpenAI, row: dict) -> dict:
    title = row.get("Title", "")
    abstract = row.get("Abstract", "")
    journal = row.get("Journal", "")
    keyword = row.get("Keyword matched", "")

    prompt = f"""
あなたは小島先生専属の論文秘書AIです。

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

以下の論文について、小島先生の研究にどれくらい関連するか判定してください。

必ずJSONだけで返してください。
余計な説明文は不要です。

形式:
{{
  "score": 0から100の整数,
  "reason": "関連度の理由を日本語で短く",
  "research_connection": "小島先生の研究との接点",
  "action": "読むべき / 要旨だけ確認 / 保留 / 無視",
  "keywords": ["関連キーワード1", "関連キーワード2", "関連キーワード3"]
}}

Journal: {journal}
Matched keyword: {keyword}
Title: {title}
Abstract: {abstract}
"""

    try:
        response = client.responses.create(
            model="gpt-5",
            input=prompt,
        )
        text = response.output_text.strip()
        text = text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        data = json.loads(text)

        return {
            "Relevance score": data.get("score", ""),
            "GPT reason": data.get("reason", ""),
            "Research connection": data.get("research_connection", ""),
            "Recommended action": data.get("action", ""),
            "GPT keywords": ", ".join(data.get("keywords", [])),
        }

    except Exception as exc:
        return {
            "Relevance score": "",
            "GPT reason": f"GPT判定エラー: {exc}",
            "Research connection": "",
            "Recommended action": "",
            "GPT keywords": "",
        }


def main():
    args = parse_args()
    load_dotenv(WORKSPACE_DIR / ".env")

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

    if not args.no_gpt:
        client = OpenAI()

        scored_rows = []
        for i, row in df.head(args.max_gpt).iterrows():
            print(f"GPT scoring {i + 1}/{min(len(df), args.max_gpt)}: {row['Title'][:60]}")
            score = score_with_gpt(client, row.to_dict())
            scored_rows.append(score)

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