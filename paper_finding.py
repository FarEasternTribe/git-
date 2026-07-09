import json
import sys
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path


WORKSPACE_DIR = Path(__file__).resolve().parent
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)
if BUNDLED_SITE_PACKAGES.exists():
    sys.path.insert(0, str(BUNDLED_SITE_PACKAGES))

import pandas as pd


JOURNALS = {
    "Nature": "0028-0836",
    "Science": "0036-8075",
    "JACS": "0002-7863",
    "Nature Chemistry": "1755-4330",
    "Nature Communications": "2041-1723",
}

KEYWORDS = [
    "graphene nanoribbon",
    "GNR",
    "nanographene",
    "STM",
    "scanning tunneling microscopy",
    "on-surface synthesis",
    "surface chemistry",
    "Au(111)",
    "molecular self-assembly",
    "electrocatalysis",
    "silicon etching",
    "chemical etching",
    "chiral",
    "spin",
    "CISS",
    "CO2RR",
    "HER",
    "OER",
]

today = date.today()
from_date = today - timedelta(days=1)
to_date = today

OUTPUT_DIR = WORKSPACE_DIR / "papers"
OUTPUT_DIR.mkdir(exist_ok=True)
output_file = OUTPUT_DIR / f"{today}_paper_list.xlsx"


def crossref_get(params: dict[str, str | int]) -> dict:
    query = urllib.parse.urlencode(params)
    url = f"https://api.crossref.org/works?{query}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "OpenAI-Agent paper_finding.py",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=12) as response:
        return json.loads(response.read().decode("utf-8"))


def get_published(item: dict) -> str:
    for key in ("published-online", "published-print", "published"):
        if key in item:
            return "-".join(map(str, item[key]["date-parts"][0]))
    return ""


rows = []
seen_dois = set()

for journal_name, issn in JOURNALS.items():
    print(f"Searching: {journal_name}")
    params = {
        "filter": f"issn:{issn},from-pub-date:{from_date},until-pub-date:{to_date},type:journal-article",
        "rows": 100,
        "select": "DOI,title,container-title,published-print,published-online,published,URL,abstract,author",
    }

    try:
        data = crossref_get(params)
        items = data["message"]["items"]
    except Exception as exc:
        print(f"Error: {journal_name}: {exc}")
        continue

    for item in items:
        doi = item.get("DOI", "")
        if doi in seen_dois:
            continue

        title = item.get("title", [""])[0]
        abstract = item.get("abstract", "")
        search_text = f"{title}\n{abstract}".lower()
        matched = [kw for kw in KEYWORDS if kw.lower() in search_text]
        if not matched:
            continue

        seen_dois.add(doi)
        authors = item.get("author", [])
        author_text = ", ".join(
            f"{author.get('given', '')} {author.get('family', '')}".strip()
            for author in authors[:5]
        )

        rows.append(
            {
                "Date found": str(today),
                "Journal": item.get("container-title", [journal_name])[0],
                "Published": get_published(item),
                "Keyword matched": ", ".join(matched),
                "Title": title,
                "Authors": author_text,
                "DOI": doi,
                "URL": item.get("URL", ""),
                "Abstract": abstract,
            }
        )

df = pd.DataFrame(rows)

if df.empty:
    print("該当論文は見つかりませんでした。")
else:
    df = df.sort_values(["Journal", "Published", "Title"])
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="papers")

    print(f"\n保存しました: {output_file}")
    print(f"件数: {len(df)}")
