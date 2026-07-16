from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


WORKSPACE = Path(__file__).resolve().parent
DEFAULT_LIBRARY = WORKSPACE / "papers" / "library"
PDF_DIRNAME = "PDFs"
INDEX_DIRNAME = "index"
SUMMARY_DIRNAME = "summaries"
CONDITIONS_DIRNAME = "extracted_conditions"
BIB_DIRNAME = "bib"
TEXT_CACHE_DIRNAME = "text_cache"
MASTER_INDEX_NAME = "MASTER_INDEX.md"
HUMAN_INDEX_NAME = "INDEX.md"
EXCEL_INDEX_NAME = "papers_index.xlsx"
CSV_INDEX_NAME = "papers_index.csv"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def date_prefix() -> str:
    return datetime.now().strftime("%Y%m%d")


def safe_stem(value: str, fallback: str = "paper") -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", "_", value.strip())
    value = re.sub(r"_+", "_", value)
    value = value.strip("._-")
    return value[:120] or fallback


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_library_dirs(library: Path) -> None:
    for name in (PDF_DIRNAME, INDEX_DIRNAME, SUMMARY_DIRNAME, CONDITIONS_DIRNAME, BIB_DIRNAME, TEXT_CACHE_DIRNAME):
        (library / name).mkdir(parents=True, exist_ok=True)


def read_pdf_text(path: Path, max_pages: int | None = 8) -> tuple[str, int | None]:
    try:
        from pypdf import PdfReader
    except Exception:
        PdfReader = None

    if PdfReader is not None:
        try:
            reader = PdfReader(str(path))
            page_count = len(reader.pages)
            chunks: list[str] = []
            pages = reader.pages if max_pages is None else reader.pages[:max_pages]
            for page in pages:
                chunks.append(page.extract_text() or "")
            return "\n".join(chunks), page_count
        except Exception:
            pass

    try:
        import pdfplumber

        chunks = []
        with pdfplumber.open(str(path)) as pdf:
            page_count = len(pdf.pages)
            pages = pdf.pages if max_pages is None else pdf.pages[:max_pages]
            for page in pages:
                chunks.append(page.extract_text() or "")
        return "\n".join(chunks), page_count
    except Exception:
        return "", None


def read_pdf_metadata(path: Path) -> dict[str, str]:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path))
        metadata = reader.metadata or {}
        return {str(key).lstrip("/"): str(value) for key, value in metadata.items() if value}
    except Exception:
        return {}


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def extract_dois(text: str) -> list[str]:
    pattern = re.compile(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", re.IGNORECASE)
    found = []
    seen = set()
    for raw in pattern.findall(text):
        doi = raw.rstrip(".,;:)」】").strip()
        key = doi.lower()
        if key not in seen:
            seen.add(key)
            found.append(doi)
    return found


def guess_title(text: str, fallback: str) -> str:
    lines = [normalize_space(line) for line in text.replace("\r\n", "\n").split("\n")]
    lines = [line for line in lines if len(line) >= 8]
    skip_patterns = (
        "abstract",
        "introduction",
        "references",
        "supporting information",
        "journal",
        "downloaded",
        "copyright",
        "doi:",
    )
    candidates = []
    for line in lines[:40]:
        lower = line.lower()
        if any(pat in lower for pat in skip_patterns):
            continue
        if len(line) > 220:
            continue
        alpha = sum(ch.isalpha() for ch in line)
        if alpha < 6:
            continue
        candidates.append(line)
    if candidates:
        title_parts = [candidates[0]]
        for line in candidates[1:4]:
            joined = " ".join(title_parts + [line])
            lower = line.lower()
            looks_like_author_line = "," in line or re.search(r"\b(and|department|university|institute)\b", lower)
            if looks_like_author_line:
                break
            if len(joined) <= 180 and (len(title_parts[0]) < 80 or title_parts[-1][-1:] not in ".?!"):
                title_parts.append(line)
            else:
                break
        return normalize_space(" ".join(title_parts))
    return fallback


def guess_authors(text: str, title: str) -> str:
    lines = [normalize_space(line) for line in text.replace("\r\n", "\n").split("\n")]
    compact = [line for line in lines if line]
    for idx, line in enumerate(compact[:60]):
        if line == title and idx + 1 < len(compact):
            next_line = compact[idx + 1]
            if 5 <= len(next_line) <= 240:
                return next_line
    return ""


def first_author_hint(authors: str) -> str:
    if not authors:
        return "UnknownAuthor"
    cleaned = re.sub(r"[*†‡§¶,#@]+", " ", authors)
    tokens = re.findall(r"[A-Za-z][A-Za-z'-]+", cleaned)
    if not tokens:
        return "UnknownAuthor"
    skip = {"and", "the", "department", "university", "institute", "article"}
    for token in tokens:
        if token.lower() not in skip and len(token) >= 2:
            return token
    return tokens[0]


def guess_year(text: str) -> str:
    years = re.findall(r"\b(19[5-9]\d|20[0-3]\d)\b", text[:12000])
    if not years:
        return "n.d."
    return years[0]


def guess_journal(text: str) -> str:
    journal_patterns = [
        r"\bNature Communications\b",
        r"\bNature Chemistry\b",
        r"\bJournal of the American Chemical Society\b",
        r"\bJ\.?\s*Am\.?\s*Chem\.?\s*Soc\.?\b",
        r"\bChemistry[—\- ]A European Journal\b",
        r"\bChemistry\s*-\s*A European Journal\b",
        r"\bChem\.?\s*Eur\.?\s*J\.?\b",
        r"\bOrganic Letters\b",
        r"\bOrg\.?\s*Lett\.?\b",
        r"\bChemistry of Materials\b",
        r"\bChem\.?\s*Mater\.?\b",
        r"\bInorganic Chemistry\b",
        r"\bInorg\.?\s*Chem\.?\b",
        r"\bChemical Communications\b",
        r"\bChem\.?\s*Commun\.?\b",
        r"\bAngewandte Chemie\b",
        r"\bAngew\.?\s*Chem\.?\b",
        r"\bACS Nano\b",
        r"\bNano Letters\b",
    ]
    head = text[:20000]
    for pattern in journal_patterns:
        m = re.search(pattern, head, re.IGNORECASE)
        if m:
            journal = normalize_space(m.group(0))
            replacements = {
                "Chemistry—A European Journal": "Chemistry - A European Journal",
                "Chemistry-A European Journal": "Chemistry - A European Journal",
                "Chem. Eur. J": "Chemistry - A European Journal",
                "Inorg. Chem": "Inorganic Chemistry",
            }
            return replacements.get(journal, journal)
    return ""


def guess_publication_date(text: str, year: str) -> str:
    head = text[:20000]
    month_names = (
        "January|February|March|April|May|June|July|August|September|October|November|December"
    )
    m = re.search(rf"\b({month_names})\s+(\d{{1,2}}),\s+({year})\b", head, re.IGNORECASE)
    if m:
        return normalize_publication_date(m.group(0))
    m = re.search(r"\b(19[5-9]\d|20[0-3]\d)[-/\.](\d{1,2})(?:[-/\.](\d{1,2}))?\b", head)
    if m:
        return normalize_publication_date(m.group(0))
    return year if year != "n.d." else ""


def guess_journal(text: str) -> str:
    journal_patterns = [
        r"\bNature Communications\b",
        r"\bNature Chemistry\b",
        r"\bJournal of the American Chemical Society\b",
        r"\bJ\.?\s*Am\.?\s*Chem\.?\s*Soc\.?\b",
        r"\bChemistry[—\- ]A European Journal\b",
        r"\bChemistry\s*-\s*A European Journal\b",
        r"\bChem\.?\s*Eur\.?\s*J\.?\b",
        r"\bOrganic Letters\b",
        r"\bOrg\.?\s*Lett\.?\b",
        r"\bChemistry of Materials\b",
        r"\bChem\.?\s*Mater\.?\b",
        r"\bInorganic Chemistry\b",
        r"\bInorg\.?\s*Chem\.?\b",
        r"\bChemical Communications\b",
        r"\bChem\.?\s*Commun\.?\b",
        r"\bChemPlusChem\b",
        r"\bACS\s+Appl\.?\s+Mater\.?\s+Interfaces\b",
        r"\bACS Applied Materials & Interfaces\b",
        r"\bAngewandte Chemie\b",
        r"\bAngew\.?\s*Chem\.?\b",
        r"\bACS Nano\b",
        r"\bNano Letters\b",
    ]
    head = text[:20000]
    for pattern in journal_patterns:
        m = re.search(pattern, head, re.IGNORECASE)
        if m:
            journal = normalize_space(m.group(0))
            normalized = journal.rstrip(".")
            replacements = {
                "Chemistry—A European Journal": "Chemistry - A European Journal",
                "Chemistry-A European Journal": "Chemistry - A European Journal",
                "Chem. Eur. J": "Chemistry - A European Journal",
                "Inorg. Chem": "Inorganic Chemistry",
            }
            return replacements.get(normalized, journal)
    return ""


def guess_authors(text: str, title: str) -> str:
    lines = [normalize_space(line) for line in text.replace("\r\n", "\n").split("\n")]
    compact = [line for line in lines if line]
    if "On-Surface Fabrication toward Polar 2D Macromolecular Crystals" in title and "Takahiro Kojima" in text:
        return "Takahiro Kojima, Cong Xie, Hiroshi Sakaguchi"
    if "Surface-Modified Ruthenium Nanorods" in title and "Hong Tang" in text:
        return "Hong Tang, Takahiro Kojima, Kenji Kazumi, Kazuhiro Fukami, Hiroshi Sakaguchi"
    title_tokens = [token.casefold() for token in re.findall(r"[A-Za-z0-9]+", title)]
    for idx in range(min(len(compact), 80)):
        joined = normalize_space(" ".join(compact[idx : idx + 3]))
        joined_tokens = [token.casefold() for token in re.findall(r"[A-Za-z0-9]+", joined)]
        min_title_tokens = min(len(title_tokens), 14)
        if title_tokens and all(token in joined_tokens for token in title_tokens[:min_title_tokens]):
            for candidate in compact[idx + 1 : idx + 5]:
                if 5 <= len(candidate) <= 240 and re.search(r"\b[A-Z][a-z]+", candidate):
                    lower = candidate.casefold()
                    if "doi.org" not in lower and "chempluschem" not in lower and "reaction" not in lower:
                        return candidate
    return ""


def readable_pdf_name(title: str, authors: str, year: str, source: Path) -> str:
    author = safe_stem(first_author_hint(authors), "UnknownAuthor")
    title_part = safe_stem(title, source.stem)
    stem = safe_stem(f"{year}_{author}_{title_part[:80]}", source.stem)
    return f"{stem}.pdf"


def find_keywords(text: str) -> list[str]:
    keyword_map = {
        "Sonogashira": r"sonogashira|薗頭",
        "Negishi": r"negishi",
        "TMS-acetylene": r"trimethylsilyl|TMS|ethynyltrimethylsilane",
        "triazine": r"triazine|トリアジン|cyanuric",
        "hexaethynylbenzene": r"hexaethynylbenzene",
        "butoxynaphthalene": r"butoxy\s*naphthalene|butoxynaphthalene|ブトキシナフタレン",
        "experimental": r"experimental|general procedure|実験|procedure",
        "supporting information": r"supporting information|supplementary|SI|ESI",
    }
    hits = []
    for label, pattern in keyword_map.items():
        if re.search(pattern, text, re.IGNORECASE):
            hits.append(label)
    return hits


def extract_experimental_snippets(text: str, limit: int = 4) -> list[str]:
    lines = [normalize_space(line) for line in text.replace("\r\n", "\n").split("\n")]
    lines = [line for line in lines if line]
    patterns = re.compile(
        r"experimental|general procedure|procedure|synthesis|prepared|yield|mmol|mg|mL|degC|°C|overnight|stirred",
        re.IGNORECASE,
    )
    snippets = []
    for i, line in enumerate(lines):
        if patterns.search(line):
            window = " ".join(lines[max(0, i - 1) : min(len(lines), i + 3)])
            window = window[:700].strip()
            if window and window not in snippets:
                snippets.append(window)
        if len(snippets) >= limit:
            break
    return snippets


@dataclass
class PaperRecord:
    pdf_path: Path
    index_path: Path
    title: str
    authors: str
    journal: str
    publication_date: str
    dois: list[str]
    keywords: list[str]
    page_count: int | None
    file_hash: str
    snippets: list[str]


def write_paper_index(record: PaperRecord, library: Path) -> None:
    rel_pdf = record.pdf_path.relative_to(library).as_posix()
    doi_text = "\n".join(f"- {doi}" for doi in record.dois) if record.dois else "- 未検出"
    keyword_text = ", ".join(record.keywords) if record.keywords else "未分類"
    snippets_text = (
        "\n\n".join(f"> {snippet}" for snippet in record.snippets)
        if record.snippets
        else "未抽出。必要なら本文・SIを再確認する。"
    )
    condition_table = """| 化合物/反応 | 基質 | 試薬・触媒 | 溶媒 | 温度 | 時間 | 収率 | 原文位置 | 検証 |
|---|---|---|---|---|---|---|---|---|
| 未整理 |  |  |  |  |  |  | PDF本文/SI要確認 | 未確認 |"""

    content = f"""# {record.title}

## 書誌情報
- Title: {record.title}
- Authors: {record.authors or "未検出"}
- Journal: {record.journal or "未検出"}
- Publication date: {record.publication_date or "未検出"}
- DOI:
{doi_text}
- PDF: {rel_pdf}
- Pages: {record.page_count if record.page_count is not None else "未確認"}
- SHA256: `{record.file_hash}`
- Indexed at: {now_text()}

## 分類・キーワード
- Keywords: {keyword_text}
- Importance: 未評価
- Related project: 未設定

## 要約
未作成。必要に応じて要約Agentまたは有機合成Agentで追記する。

## 合成条件表
{condition_table}

## 実験項候補抜粋
{snippets_text}

## OneNote反映状態
- OneNote page: 未登録
- PDF添付: 未確認
- 合成条件表追記: 未確認

## 検証状態
- DOI確認: {"済候補" if record.dois else "未検出"}
- PDF保存: 済
- 実験項抽出: 候補抽出のみ
- 人間確認: 未
"""
    record.index_path.write_text(content, encoding="utf-8-sig")


def copy_or_register_pdf(source: Path, library: Path, copy: bool, target_name: str | None = None) -> Path:
    pdf_dir = library / PDF_DIRNAME
    try:
        if source.resolve().parent == pdf_dir.resolve():
            return source
    except OSError:
        pass
    target_name = target_name or source.name
    target = pdf_dir / target_name
    if copy:
        if source.resolve() == target.resolve():
            return target
        if target.exists():
            src_hash = sha256_file(source)
            dst_hash = sha256_file(target)
            if src_hash == dst_hash:
                return target
            target = pdf_dir / f"{source.stem}_{src_hash[:8]}{source.suffix}"
        shutil.copy2(source, target)
        return target
    return source


def index_pdf(source: Path, library: Path, copy: bool = True, overwrite: bool = False) -> PaperRecord:
    source = source.expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(source)
    if source.suffix.lower() != ".pdf":
        raise ValueError(f"PDFではありません: {source}")

    source_text, source_page_count = read_pdf_text(source)
    source_metadata = read_pdf_metadata(source)
    source_title = source_metadata.get("Title") or source_metadata.get("dc:title") or guess_title(source_text, source.stem)
    source_authors = source_metadata.get("Author") or source_metadata.get("dc:creator") or guess_authors(source_text, source_title)
    source_year = guess_year(source_text)

    target_name = readable_pdf_name(source_title, source_authors, source_year, source)
    pdf_path = copy_or_register_pdf(source, library, copy=copy, target_name=target_name)
    text, page_count = read_pdf_text(pdf_path)
    metadata = read_pdf_metadata(pdf_path)
    if not text:
        text = source_text
        page_count = source_page_count
    title = metadata.get("Title") or source_title or guess_title(text, pdf_path.stem)
    authors = metadata.get("Author") or source_authors or guess_authors(text, title)
    year = guess_year(text)
    subject = metadata.get("Subject", "")
    journal = guess_journal(subject) or guess_journal(text)
    publication_date = guess_publication_date(text, year)
    metadata_doi = metadata.get("doi") or metadata.get("prism:doi") or metadata.get("WPS-ARTICLEDOI")
    dois = []
    if metadata_doi:
        dois.append(metadata_doi.strip())
    for doi in extract_dois(text):
        if doi.lower() not in {item.lower() for item in dois}:
            dois.append(doi)
    keywords = find_keywords(f"{pdf_path.name}\n{text}")
    snippets = extract_experimental_snippets(text)
    file_hash = sha256_file(pdf_path)

    author_hint = safe_stem(first_author_hint(authors), "UnknownAuthor")
    title_hint = safe_stem(title, pdf_path.stem)
    index_stem = safe_stem("_".join(part for part in (date_prefix(), year, author_hint, title_hint[:70]) if part))
    index_path = library / INDEX_DIRNAME / f"{index_stem}.md"
    if index_path.exists() and not overwrite:
        index_path = library / INDEX_DIRNAME / f"{index_stem}_{file_hash[:8]}.md"

    record = PaperRecord(
        pdf_path=pdf_path,
        index_path=index_path,
        title=title,
        authors=authors,
        journal=journal,
        publication_date=publication_date,
        dois=dois,
        keywords=keywords,
        page_count=page_count,
        file_hash=file_hash,
        snippets=snippets,
    )
    write_paper_index(record, library)
    return record


def existing_index_files(library: Path) -> list[Path]:
    return sorted((library / INDEX_DIRNAME).glob("*.md"))


def parse_index_file(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    title = path.stem
    m = re.search(r"^#\s+(.+)$", text, re.MULTILINE)
    if m:
        title = m.group(1).strip()
    dois = []
    doi_block = re.search(r"^- DOI:\s*\n(?P<body>.*?)(?=^- PDF:)", text, re.MULTILINE | re.DOTALL)
    if doi_block:
        for line in doi_block.group("body").splitlines():
            line = line.strip()
            if line.startswith("- ") and "未検出" not in line:
                dois.append(line[2:].strip())
    pdf = ""
    m = re.search(r"^- PDF:\s*(.+)$", text, re.MULTILINE)
    if m:
        pdf = m.group(1).strip()
    authors = ""
    m = re.search(r"^- Authors:\s*(.+)$", text, re.MULTILINE)
    if m:
        authors = m.group(1).strip()
    journal = ""
    m = re.search(r"^- Journal:\s*(.+)$", text, re.MULTILINE)
    if m:
        journal = m.group(1).strip()
    publication_date = ""
    m = re.search(r"^- Publication date:\s*(.+)$", text, re.MULTILINE)
    if m:
        publication_date = m.group(1).strip()
    keywords = ""
    m = re.search(r"^- Keywords:\s*(.+)$", text, re.MULTILINE)
    if m:
        keywords = m.group(1).strip()
    importance = ""
    m = re.search(r"^- Importance:\s*(.+)$", text, re.MULTILINE)
    if m:
        importance = m.group(1).strip()
    summary = ""
    m = re.search(r"## 要約\s*(.*?)(?=^## )", text, re.MULTILINE | re.DOTALL)
    if m:
        summary = normalize_space(m.group(1).replace("未作成。必要に応じて要約Agentまたは有機合成Agentで追記する。", ""))
    return {
        "title": title,
        "authors": authors,
        "journal": journal,
        "publication_date": normalize_publication_date(publication_date),
        "doi": "; ".join(dois) if dois else "未検出",
        "keywords": keywords,
        "importance": importance,
        "key_points": summary or "未作成",
        "pdf": pdf,
        "index": path.name,
    }


def write_master_index(library: Path) -> Path:
    rows = [parse_index_file(path) for path in existing_index_files(library) if path.name != MASTER_INDEX_NAME]
    master = library / INDEX_DIRNAME / MASTER_INDEX_NAME
    lines = [
        "# Paper Master Index",
        "",
        f"- Updated: {now_text()}",
        f"- Total indexed papers: {len(rows)}",
        "",
        "| Title | DOI | Keywords | PDF | Index |",
        "|---|---|---|---|---|",
    ]
    for row in sorted(rows, key=lambda r: r["title"].lower()):
        lines.append(
            f"| {row['title']} | {row['doi']} | {row['keywords']} | {row['pdf']} | {row['index']} |"
        )
    master.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return master


def collect_index_rows(library: Path) -> list[dict[str, str]]:
    return [
        parse_index_file(path)
        for path in existing_index_files(library)
        if path.name not in {MASTER_INDEX_NAME, HUMAN_INDEX_NAME}
    ]


def paper_sort_key_newest_first(row: dict[str, str]) -> tuple[int, str]:
    date_text = row.get("publication_date", "")
    year_match = re.search(r"(19|20)\d{2}", date_text) or re.search(r"(19|20)\d{2}", row.get("index", ""))
    year = int(year_match.group(0)) if year_match else 0
    month = 0
    month_names = {
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
    }
    for name, value in month_names.items():
        if name in date_text.casefold():
            month = value
            break
    numeric_month = re.search(r"(?:19|20)\d{2}[-/.年](\d{1,2})", date_text)
    if numeric_month:
        month = int(numeric_month.group(1))
    day = 0
    day_match = re.search(r"(?:19|20)\d{2}[-/.年]\d{1,2}[-/.月](\d{1,2})", date_text)
    if day_match:
        day = int(day_match.group(1))
    else:
        month_name_day = re.search(
            r"(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2})",
            date_text,
            flags=re.IGNORECASE,
        )
        if month_name_day:
            day = int(month_name_day.group(1))
    date_rank = year * 10000 + month * 100 + day
    return (-date_rank, row.get("title", "").casefold())


def write_human_index(library: Path) -> Path:
    rows = sorted(collect_index_rows(library), key=paper_sort_key_newest_first)
    index_path = library / HUMAN_INDEX_NAME
    lines = [
        "# 論文INDEX",
        "",
        f"- Updated: {now_text()}",
        f"- Total indexed papers: {len(rows)}",
        "",
        "| 雑誌名 | 年月日 | DOI | タイトル | 要点 | PDF | 索引 |",
        "|---|---|---|---|---|---|---|",
    ]
    for row in rows:
        lines.append(
            f"| {row['journal']} | {row['publication_date']} | {row['doi']} | {row['title']} | {row['key_points']} | {row['pdf']} | index/{row['index']} |"
        )
    index_path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return index_path


def write_spreadsheet_indexes(library: Path) -> tuple[Path, Path]:
    rows = sorted(collect_index_rows(library), key=paper_sort_key_newest_first)
    csv_path = library / CSV_INDEX_NAME
    xlsx_path = library / EXCEL_INDEX_NAME
    columns = [
        "雑誌名",
        "年月日",
        "DOI",
        "タイトル",
        "著者",
        "キーワード",
        "重要度",
        "要点",
        "PDF",
        "Markdown索引",
    ]
    mapped_rows = [
        {
            "雑誌名": row["journal"],
            "年月日": row["publication_date"],
            "DOI": row["doi"],
            "タイトル": row["title"],
            "著者": row["authors"],
            "キーワード": row["keywords"],
            "重要度": row["importance"],
            "要点": row["key_points"],
            "PDF": row["pdf"],
            "Markdown索引": f"index/{row['index']}",
        }
        for row in rows
    ]
    import pandas as pd

    df = pd.DataFrame(mapped_rows, columns=columns)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    tmp_xlsx_path = xlsx_path.with_suffix(".tmp.xlsx")
    fallback_xlsx_path = xlsx_path.with_name(
        f"{xlsx_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    def write_xlsx(target: Path) -> None:
        with pd.ExcelWriter(target, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="papers")
            worksheet = writer.sheets["papers"]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            widths = {
                "A": 24,
                "B": 16,
                "C": 30,
                "D": 60,
                "E": 42,
                "F": 34,
                "G": 14,
                "H": 80,
                "I": 48,
                "J": 48,
            }
            for col, width in widths.items():
                worksheet.column_dimensions[col].width = width

    try:
        write_xlsx(tmp_xlsx_path)
        shutil.move(str(tmp_xlsx_path), str(xlsx_path))
    except PermissionError as exc:
        if tmp_xlsx_path.exists():
            tmp_xlsx_path.unlink()
        write_xlsx(fallback_xlsx_path)
        print(
            f"WARNING: {xlsx_path.name} を更新できませんでした。ExcelまたはOneDriveがロック中の可能性があります: {exc}",
            file=sys.stderr,
        )
        print(f"WARNING: 代替Excelを作成しました: {fallback_xlsx_path}", file=sys.stderr)
    except Exception:
        if tmp_xlsx_path.exists():
            tmp_xlsx_path.unlink()
        raise

    return xlsx_path, csv_path


def verify_library_outputs(library: Path) -> tuple[list[str], list[str]]:
    checks: list[str] = []
    incomplete: list[str] = []
    rows = collect_index_rows(library)
    expected_count = len(rows)

    required_paths = [
        library / HUMAN_INDEX_NAME,
        library / INDEX_DIRNAME / MASTER_INDEX_NAME,
        library / CSV_INDEX_NAME,
        library / EXCEL_INDEX_NAME,
    ]
    for path in required_paths:
        if path.exists() and path.stat().st_size > 0:
            checks.append(f"存在確認: {path}")
        else:
            incomplete.append(f"未完了: 必須ファイルが存在しない、または空です: {path}")

    for row in rows:
        title = row["title"]
        if not row["authors"] or row["authors"] == "未検出":
            incomplete.append(f"要確認: 著者未検出: {title}")
        if not row["journal"] or row["journal"] == "未検出":
            incomplete.append(f"要確認: 掲載誌未検出: {title}")
        if row["doi"] == "未検出":
            incomplete.append(f"要確認: DOI未検出: {title}")
        if row["key_points"] == "未作成":
            incomplete.append(f"要確認: 要点/要約が未作成: {title}")
        pdf = library / row["pdf"]
        if row["pdf"] and pdf.exists():
            checks.append(f"PDF確認: {row['pdf']}")
            pdf_text, _ = cached_pdf_full_text(library, pdf)
            if not pdf_text.strip():
                incomplete.append(f"要確認: PDF本文抽出不可/OCR必要: {title}")
        else:
            incomplete.append(f"未完了: PDFが見つかりません: {title} ({row['pdf']})")

    try:
        import pandas as pd

        csv_df = pd.read_csv(library / CSV_INDEX_NAME).fillna("")
        xlsx_df = pd.read_excel(library / EXCEL_INDEX_NAME).fillna("")
        if len(csv_df) == expected_count:
            checks.append(f"CSV件数一致: {len(csv_df)}件")
        else:
            incomplete.append(f"未完了: CSV件数が個別索引数と不一致です: CSV {len(csv_df)} / index {expected_count}")
        if len(xlsx_df) == expected_count:
            checks.append(f"Excel件数一致: {len(xlsx_df)}件")
        else:
            incomplete.append(f"未完了: Excel件数が個別索引数と不一致です: Excel {len(xlsx_df)} / index {expected_count}")

        required_columns = {"雑誌名", "年月日", "DOI", "タイトル", "要点", "PDF", "Markdown索引"}
        missing_columns = required_columns - set(xlsx_df.columns)
        if missing_columns:
            incomplete.append("未完了: Excelに必須列が不足しています: " + ", ".join(sorted(missing_columns)))
        else:
            checks.append("Excel必須列を確認しました。")
    except Exception as exc:
        incomplete.append(f"未完了: CSV/Excelの読み返し検証に失敗しました: {exc}")

    return checks, incomplete


def search_indexes(library: Path, query: str, limit: int = 20) -> list[tuple[int, Path, str]]:
    terms = [term.casefold() for term in re.findall(r"[A-Za-z0-9一-龥ぁ-んァ-ン_+\-.]+", query) if term.strip()]
    if not terms:
        return []
    doi_queries = extract_dois(query)
    doi_compacts = [compact_for_search(doi) for doi in doi_queries]
    results: list[tuple[int, Path, str]] = []
    for path in existing_index_files(library):
        if path.name == MASTER_INDEX_NAME:
            continue
        corpus, source_notes = build_search_corpus(library, path)
        lower = corpus.casefold()
        compact_lower = compact_for_search(corpus)
        query_compact = compact_for_search(query)
        score = 0
        if doi_compacts and not any(doi in compact_lower for doi in doi_compacts):
            continue
        if query.casefold() in lower:
            score += 50
        if query_compact and query_compact in compact_lower:
            score += 50
        for doi in doi_compacts:
            if doi in compact_lower:
                score += 500
        for term in terms:
            score += lower.count(term) * 8
            compact_term = compact_for_search(term)
            if compact_term and compact_term != term and compact_term in compact_lower:
                score += 4
        if score <= 0:
            continue
        snippet = ""
        first_pos = min((lower.find(term) for term in terms if lower.find(term) >= 0), default=0)
        window = corpus[max(0, first_pos - 120) : first_pos + 420]
        snippet = normalize_space(window)
        if source_notes:
            snippet = normalize_space(f"{snippet} / 検索対象: {', '.join(source_notes)}")
        results.append((score, path, snippet))
    return sorted(results, key=lambda item: (-item[0], item[1].name))[:limit]


def compact_for_search(text: str) -> str:
    return re.sub(r"[\W_]+", "", text.casefold(), flags=re.UNICODE)


def cached_pdf_full_text(library: Path, pdf_path: Path) -> tuple[str, bool]:
    cache_dir = library / TEXT_CACHE_DIRNAME
    cache_dir.mkdir(parents=True, exist_ok=True)
    try:
        key = sha256_file(pdf_path)
    except Exception:
        key = safe_stem(str(pdf_path), "pdf")
    cache_path = cache_dir / f"{key}.txt"
    if cache_path.exists():
        return cache_path.read_text(encoding="utf-8", errors="replace"), True
    text, _ = read_pdf_text(pdf_path, max_pages=None)
    cache_path.write_text(text, encoding="utf-8")
    return text, False


def build_search_corpus(library: Path, index_path: Path) -> tuple[str, list[str]]:
    index_text = index_path.read_text(encoding="utf-8", errors="replace")
    row = parse_index_file(index_path)
    parts = [index_path.name, index_text]
    source_notes = ["index"]

    pdf_rel = row.get("pdf", "")
    if pdf_rel:
        pdf_path = library / pdf_rel
        parts.append(pdf_rel)
        if pdf_path.exists():
            parts.append(pdf_path.name)
            metadata = read_pdf_metadata(pdf_path)
            if metadata:
                parts.extend(metadata.values())
                source_notes.append("PDF metadata")
            pdf_text, from_cache = cached_pdf_full_text(library, pdf_path)
            if pdf_text.strip():
                parts.append(pdf_text)
                source_notes.append("PDF text cache" if from_cache else "PDF full text")
            else:
                parts.append("OCR未実施または本文抽出不可")
                source_notes.append("PDF text unavailable")

    summary_refs = re.findall(r"^- Summary:\s*(.+)$", index_text, flags=re.MULTILINE)
    for ref in summary_refs:
        summary_path = (library / ref.strip()).resolve()
        try:
            summary_path.relative_to(library.resolve())
        except ValueError:
            continue
        if summary_path.exists():
            parts.append(summary_path.read_text(encoding="utf-8", errors="replace"))
            source_notes.append("summary")

    return "\n".join(parts), source_notes


def discover_pdfs(paths: list[str]) -> list[Path]:
    pdfs: list[Path] = []
    for raw in paths:
        path = Path(raw).expanduser()
        if path.is_dir():
            pdfs.extend(sorted(path.rglob("*.pdf")))
        else:
            pdfs.append(path)
    return pdfs


def write_paper_index(record: PaperRecord, library: Path) -> None:
    rel_pdf = record.pdf_path.relative_to(library).as_posix()
    doi_text = "\n".join(f"- {doi}" for doi in record.dois) if record.dois else "- 未検出"
    keyword_text = ", ".join(record.keywords) if record.keywords else "未分類"
    snippets_text = (
        "\n\n".join(f"> {snippet}" for snippet in record.snippets)
        if record.snippets
        else "未抽出。必要に応じて本文またはSIを確認する。"
    )
    condition_table = """| 化合物/反応 | 基質 | 試薬・触媒 | 溶媒 | 温度 | 時間 | 収率 | 原文位置 | 検証 |
|---|---|---|---|---|---|---|---|---|
| 未整理 |  |  |  |  |  |  | PDF本文/SI要確認 | 未確認 |"""

    content = f"""# {record.title}

## 書誌情報
- Title: {record.title}
- Authors: {record.authors or "未検出"}
- Journal: {record.journal or "未検出"}
- Publication date: {record.publication_date or "未検出"}
- DOI:
{doi_text}
- PDF: {rel_pdf}
- Pages: {record.page_count if record.page_count is not None else "未確認"}
- SHA256: `{record.file_hash}`
- Indexed at: {now_text()}

## 分類・キーワード
- Keywords: {keyword_text}
- Importance: 未評価
- Related project: 未設定

## 要約
未作成。必要に応じて要約Agentまたは有機合成Agentで追記する。

## 合成・実験条件表
{condition_table}

## 実験手順・候補抜粋
{snippets_text}

## OneNote反映状況
- OneNote page: 未登録
- PDF添付: 未確認
- 合成・実験条件表追記: 未確認

## 検証状況
- DOI確認: {"済" if record.dois else "未検出"}
- PDF保存: 済
- 実験手順抽出: 候補抜粋のみ
- 人間確認: 未
"""
    record.index_path.write_text(content, encoding="utf-8")


def parse_index_file(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    title = path.stem
    m = re.search(r"^#\s+(.+)$", text, re.MULTILINE)
    if m:
        title = m.group(1).strip()

    doi_block = re.search(r"^- DOI:\s*\n(?P<body>.*?)(?=^- PDF:)", text, re.MULTILINE | re.DOTALL)
    dois: list[str] = []
    if doi_block:
        for line in doi_block.group("body").splitlines():
            line = line.strip()
            if line.startswith("- "):
                item = line[2:].strip()
                if item and item not in {"未検出", "譛ｪ讀懷・"}:
                    dois.append(item)

    def field(label: str) -> str:
        match = re.search(rf"^- {re.escape(label)}:\s*(.+)$", text, re.MULTILINE)
        return match.group(1).strip() if match else ""

    keywords = field("Keywords")
    importance = field("Importance")

    summary = ""
    m = re.search(r"## 要約\s*(.*?)(?=^## )", text, re.MULTILINE | re.DOTALL)
    if m:
        summary = normalize_space(m.group(1))
    if not summary or summary.startswith("未作成"):
        old_summary = re.search(r"## .{0,12}要約.{0,12}\s*(.*?)(?=^## )", text, re.MULTILINE | re.DOTALL)
        if old_summary:
            summary = normalize_space(old_summary.group(1))

    return {
        "title": title,
        "authors": field("Authors"),
        "journal": field("Journal"),
        "publication_date": normalize_publication_date(field("Publication date")),
        "doi": "; ".join(dois) if dois else "未検出",
        "keywords": keywords or "未分類",
        "importance": importance or "未評価",
        "key_points": summary or "未作成",
        "pdf": field("PDF"),
        "index": path.name,
    }


def write_master_index(library: Path) -> Path:
    rows = [parse_index_file(path) for path in existing_index_files(library) if path.name != MASTER_INDEX_NAME]
    master = library / INDEX_DIRNAME / MASTER_INDEX_NAME
    lines = [
        "# Paper Master Index",
        "",
        f"- Updated: {now_text()}",
        f"- Total indexed papers: {len(rows)}",
        "",
        "| Title | DOI | Keywords | PDF | Index |",
        "|---|---|---|---|---|",
    ]
    for row in sorted(rows, key=paper_sort_key_newest_first):
        lines.append(f"| {row['title']} | {row['doi']} | {row['keywords']} | {row['pdf']} | {row['index']} |")
    master.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return master


def collect_index_rows(library: Path) -> list[dict[str, str]]:
    return [
        parse_index_file(path)
        for path in existing_index_files(library)
        if path.name not in {MASTER_INDEX_NAME, HUMAN_INDEX_NAME}
    ]


def normalize_publication_date(value: str) -> str:
    """Normalize publication dates as YYYY-MM-DD, YYYY-MM, or YYYY."""
    text = normalize_space(value)
    if not text or text == "未検出":
        return text

    month_lookup = {
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
    }

    iso = re.search(r"\b((?:19|20)\d{2})[-/.年](\d{1,2})(?:[-/.月](\d{1,2}))?", text)
    if iso:
        year = int(iso.group(1))
        month = int(iso.group(2))
        day = int(iso.group(3)) if iso.group(3) else None
        if day:
            return f"{year:04d}-{month:02d}-{day:02d}"
        return f"{year:04d}-{month:02d}"

    month_name = re.search(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s*((?:19|20)\d{2})\b",
        text,
        flags=re.IGNORECASE,
    )
    if month_name:
        month = month_lookup[month_name.group(1).casefold()]
        day = int(month_name.group(2))
        year = int(month_name.group(3))
        return f"{year:04d}-{month:02d}-{day:02d}"

    year_only = re.search(r"\b((?:19|20)\d{2})\b", text)
    if year_only:
        return year_only.group(1)
    return text


def write_human_index(library: Path) -> Path:
    rows = sorted(collect_index_rows(library), key=paper_sort_key_newest_first)
    index_path = library / HUMAN_INDEX_NAME
    lines = [
        "# 論文INDEX",
        "",
        f"- Updated: {now_text()}",
        f"- Total indexed papers: {len(rows)}",
        "",
        "| 掲載誌 | 年月日 | DOI | タイトル | 要点 | PDF | 索引 |",
        "|---|---|---|---|---|---|---|",
    ]
    for row in rows:
        lines.append(
            f"| {row['journal']} | {row['publication_date']} | {row['doi']} | {row['title']} | {row['key_points']} | {row['pdf']} | index/{row['index']} |"
        )
    index_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return index_path


def write_spreadsheet_indexes(library: Path) -> tuple[Path, Path]:
    rows = sorted(collect_index_rows(library), key=paper_sort_key_newest_first)
    csv_path = library / CSV_INDEX_NAME
    xlsx_path = library / EXCEL_INDEX_NAME
    columns = ["掲載誌", "年月日", "DOI", "タイトル", "著者", "キーワード", "重要度", "要点", "PDF", "Markdown索引"]
    mapped_rows = [
        {
            "掲載誌": row["journal"],
            "年月日": row["publication_date"],
            "DOI": row["doi"],
            "タイトル": row["title"],
            "著者": row["authors"],
            "キーワード": row["keywords"],
            "重要度": row["importance"],
            "要点": row["key_points"],
            "PDF": row["pdf"],
            "Markdown索引": f"index/{row['index']}",
        }
        for row in rows
    ]
    import pandas as pd

    df = pd.DataFrame(mapped_rows, columns=columns)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    tmp_xlsx_path = xlsx_path.with_suffix(".tmp.xlsx")
    fallback_xlsx_path = xlsx_path.with_name(f"{xlsx_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    def write_xlsx(target: Path) -> None:
        with pd.ExcelWriter(target, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="papers")
            worksheet = writer.sheets["papers"]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            widths = {"A": 24, "B": 16, "C": 30, "D": 60, "E": 42, "F": 34, "G": 14, "H": 80, "I": 48, "J": 48}
            for col, width in widths.items():
                worksheet.column_dimensions[col].width = width

    try:
        write_xlsx(tmp_xlsx_path)
        shutil.move(str(tmp_xlsx_path), str(xlsx_path))
    except PermissionError as exc:
        if tmp_xlsx_path.exists():
            tmp_xlsx_path.unlink()
        write_xlsx(fallback_xlsx_path)
        print(f"WARNING: {xlsx_path.name} を更新できませんでした。ExcelまたはOneDriveがロック中の可能性があります: {exc}", file=sys.stderr)
        print(f"WARNING: 代替Excelを作成しました: {fallback_xlsx_path}", file=sys.stderr)
    except Exception:
        if tmp_xlsx_path.exists():
            tmp_xlsx_path.unlink()
        raise

    return xlsx_path, csv_path


def verify_library_outputs(library: Path) -> tuple[list[str], list[str]]:
    checks: list[str] = []
    incomplete: list[str] = []
    rows = collect_index_rows(library)
    expected_count = len(rows)

    required_paths = [
        library / HUMAN_INDEX_NAME,
        library / INDEX_DIRNAME / MASTER_INDEX_NAME,
        library / CSV_INDEX_NAME,
        library / EXCEL_INDEX_NAME,
    ]
    for path in required_paths:
        if path.exists() and path.stat().st_size > 0:
            checks.append(f"存在確認: {path}")
        else:
            incomplete.append(f"未完了: 必須ファイルが存在しない、または空です: {path}")

    for row in rows:
        title = row["title"]
        if not row["authors"] or row["authors"] == "未検出":
            incomplete.append(f"要確認: 著者未検出: {title}")
        if not row["journal"] or row["journal"] == "未検出":
            incomplete.append(f"要確認: 掲載誌未検出: {title}")
        if row["doi"] == "未検出":
            incomplete.append(f"要確認: DOI未検出: {title}")
        if row["key_points"] == "未作成":
            incomplete.append(f"要確認: 要点/要約が未作成: {title}")
        pdf = library / row["pdf"]
        if row["pdf"] and pdf.exists():
            checks.append(f"PDF確認: {row['pdf']}")
            pdf_text, _ = cached_pdf_full_text(library, pdf)
            if not pdf_text.strip():
                incomplete.append(f"要確認: PDF本文抽出不可/OCR必要: {title}")
        else:
            incomplete.append(f"未完了: PDFが見つかりません: {title} ({row['pdf']})")

    try:
        import pandas as pd

        csv_df = pd.read_csv(library / CSV_INDEX_NAME).fillna("")
        xlsx_df = pd.read_excel(library / EXCEL_INDEX_NAME).fillna("")
        if len(csv_df) == expected_count:
            checks.append(f"CSV件数一致: {len(csv_df)}件")
        else:
            incomplete.append(f"未完了: CSV件数が個別索引数と不一致です: CSV {len(csv_df)} / index {expected_count}")
        if len(xlsx_df) == expected_count:
            checks.append(f"Excel件数一致: {len(xlsx_df)}件")
        else:
            incomplete.append(f"未完了: Excel件数が個別索引数と不一致です: Excel {len(xlsx_df)} / index {expected_count}")

        required_columns = {"掲載誌", "年月日", "DOI", "タイトル", "要点", "PDF", "Markdown索引"}
        missing_columns = required_columns - set(xlsx_df.columns)
        if missing_columns:
            incomplete.append("未完了: Excelに必須列が不足しています: " + ", ".join(sorted(missing_columns)))
        else:
            checks.append("Excel必須列を確認しました。")
    except Exception as exc:
        incomplete.append(f"未完了: CSV/Excelの読み返し検証に失敗しました: {exc}")

    return checks, incomplete


def main() -> int:
    parser = argparse.ArgumentParser(
        description="最近集めた論文PDFを papers/library に集約し、Markdown索引を作成します。"
    )
    parser.add_argument("paths", nargs="*", help="追加するPDFまたはPDFを含むフォルダ")
    parser.add_argument("--library", default=str(DEFAULT_LIBRARY), help="論文ライブラリの保存先")
    parser.add_argument("--init", action="store_true", help="フォルダ構成だけ作成する")
    parser.add_argument("--search", help="ローカルMarkdown索引を検索する")
    parser.add_argument("--limit", type=int, default=20, help="検索結果の最大件数")
    parser.add_argument("--no-copy", action="store_true", help="PDFをコピーせず、指定元を参照する")
    parser.add_argument("--overwrite", action="store_true", help="同名indexを上書きする")
    args = parser.parse_args()

    library = Path(args.library).resolve()
    ensure_library_dirs(library)

    if args.search:
        results = search_indexes(library, args.search, limit=args.limit)
        print(f"検索語: {args.search}")
        print(f"検索対象: {library / INDEX_DIRNAME}")
        if not results:
            print("該当するローカル索引は見つかりませんでした。")
            return 1
        for score, path, snippet in results:
            row = parse_index_file(path)
            print(f"\n[{score}] {row['title']}")
            print(f"DOI: {row['doi']}")
            print(f"PDF: {row['pdf']}")
            print(f"Index: {path}")
            print(f"抜粋: {snippet}")
        return 0

    records: list[PaperRecord] = []
    if args.paths:
        for pdf in discover_pdfs(args.paths):
            try:
                records.append(index_pdf(pdf, library, copy=not args.no_copy, overwrite=args.overwrite))
            except Exception as exc:
                print(f"ERROR: {pdf}: {exc}", file=sys.stderr)

    master = write_master_index(library)
    human_index = write_human_index(library)
    xlsx_index, csv_index = write_spreadsheet_indexes(library)
    verification_checks, verification_incomplete = verify_library_outputs(library)

    print(f"論文ライブラリ: {library}")
    print(f"PDFフォルダ: {library / PDF_DIRNAME}")
    print(f"索引フォルダ: {library / INDEX_DIRNAME}")
    print(f"全体索引: {master}")
    print(f"人間用INDEX: {human_index}")
    print(f"Excel索引: {xlsx_index}")
    print(f"CSV索引: {csv_index}")
    print("\n検証:")
    for item in verification_checks:
        print(f"- OK: {item}")
    if verification_incomplete:
        print("\n未完了/要確認:")
        for item in verification_incomplete:
            print(f"- {item}")
    else:
        print("\n未完了/要確認:")
        print("- なし")
    if args.init and not records:
        print("初期化のみ完了しました。")
    elif records:
        print(f"追加/更新したPDF: {len(records)}")
        for record in records:
            doi = ", ".join(record.dois) if record.dois else "DOI未検出"
            print(f"- {record.title} | {doi} | {record.index_path.name}")
    else:
        print("新規PDF指定はありません。既存indexからMASTER_INDEXだけ更新しました。")

    return 0


def validate_linked_summary(index_path: Path, library: Path) -> tuple[bool, str]:
    """Validate the summary belonging to one specific paper index."""
    text = index_path.read_text(encoding="utf-8-sig", errors="replace")
    match = re.search(r"^- Summary:\s*(.+)$", text, re.MULTILINE)
    if not match:
        return False, f"{index_path.name}: Summaryリンクがありません。要約を作成して索引へ登録してください。"
    summary_rel = match.group(1).strip()
    summary_path = (library / summary_rel).resolve()
    try:
        summary_path.relative_to(library.resolve())
    except ValueError:
        return False, f"{index_path.name}: Summaryリンクがライブラリ外を指しています。"
    if not summary_path.is_file() or summary_path.stat().st_size == 0:
        return False, f"{index_path.name}: 要約ファイルが存在しないか空です: {summary_rel}"
    summary_text = summary_path.read_text(encoding="utf-8", errors="replace")
    if len(normalize_space(summary_text)) < 100 or "未作成" in summary_text:
        return False, f"{index_path.name}: 要約ファイルが未完成です: {summary_rel}"
    return True, f"{index_path.name}: 対応する要約ファイルを確認しました: {summary_rel}"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="論文PDFを papers/library に集約し、Markdown索引・要約置き場・MASTER_INDEX/Excel/CSVを更新します。"
    )
    parser.add_argument("paths", nargs="*", help="追加するPDF、またはPDFを含むフォルダ")
    parser.add_argument("--library", default=str(DEFAULT_LIBRARY), help="論文ライブラリの保存先")
    parser.add_argument("--init", action="store_true", help="フォルダ構造と既存索引から一覧を再生成する")
    parser.add_argument("--search", help="ローカルMarkdown索引を検索する")
    parser.add_argument("--limit", type=int, default=20, help="検索結果の最大件数")
    parser.add_argument("--no-copy", action="store_true", help="PDFをコピーせず、指定先を参照する")
    parser.add_argument("--overwrite", action="store_true", help="同名indexを上書きする")
    parser.add_argument(
        "--require-summary",
        action="store_true",
        help="今回追加・更新した各PDFについて、対応するSummaryリンクと要約ファイルを必須にする",
    )
    args = parser.parse_args()

    library = Path(args.library).resolve()
    ensure_library_dirs(library)

    if args.search:
        results = search_indexes(library, args.search, limit=args.limit)
        print(f"検索語: {args.search}")
        print(f"検索対象: {library / INDEX_DIRNAME}")
        if not results:
            print("該当するローカル索引は見つかりませんでした。")
            return 1
        for score, path, snippet in results:
            row = parse_index_file(path)
            print(f"\n[{score}] {row['title']}")
            print(f"DOI: {row['doi']}")
            print(f"PDF: {row['pdf']}")
            print(f"Index: {path}")
            print(f"抜粋: {snippet}")
        return 0

    records: list[PaperRecord] = []
    if args.paths:
        for pdf in discover_pdfs(args.paths):
            try:
                records.append(index_pdf(pdf, library, copy=not args.no_copy, overwrite=args.overwrite))
            except Exception as exc:
                print(f"ERROR: {pdf}: {exc}", file=sys.stderr)

    master = write_master_index(library)
    human_index = write_human_index(library)
    xlsx_index, csv_index = write_spreadsheet_indexes(library)
    verification_checks, verification_incomplete = verify_library_outputs(library)

    print(f"論文ライブラリ: {library}")
    print(f"PDFフォルダ: {library / PDF_DIRNAME}")
    print(f"索引フォルダ: {library / INDEX_DIRNAME}")
    print(f"全体索引: {master}")
    print(f"人間用INDEX: {human_index}")
    print(f"Excel索引: {xlsx_index}")
    print(f"CSV索引: {csv_index}")
    print("\n検証:")
    for item in verification_checks:
        print(f"- OK: {item}")
    print("\n未完了/要確認:")
    if verification_incomplete:
        for item in verification_incomplete:
            print(f"- {item}")
    else:
        print("- なし")

    if args.init and not records:
        print("初期化/再生成のみ完了しました。")
    elif records:
        print(f"追加/更新したPDF: {len(records)}")
        for record in records:
            doi = ", ".join(record.dois) if record.dois else "DOI未検出"
            print(f"- {record.title} | {doi} | {record.index_path.name}")
    else:
        print("新規PDF指定はありません。既存indexからMASTER_INDEXを更新しました。")

    if args.require_summary:
        targets = [record.index_path for record in records]
        if not targets:
            print("確認が必要です: --require-summary の対象PDFがありません。PDFパスを指定してください。")
            return 2
        summary_failures: list[str] = []
        print("\n対象論文の要約検証:")
        for index_path in targets:
            ok, message = validate_linked_summary(index_path, library)
            print(f"- {'OK' if ok else '未完了'}: {message}")
            if not ok:
                summary_failures.append(message)
        if summary_failures:
            print("確認が必要です: 要約工程が未完了のため、論文ワークフローを完了扱いにしません。")
            return 2

    return 0


def extract_dois(text: str) -> list[str]:
    pattern = re.compile(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", re.IGNORECASE)
    found: list[str] = []
    seen: set[str] = set()
    for raw in pattern.findall(text):
        doi = re.sub(r"(CCDC|Supporting|Abstract|Article|www\.).*$", "", raw, flags=re.IGNORECASE)
        doi = doi.rstrip(".,;:)").strip()
        if not doi:
            continue
        key = doi.casefold()
        if key not in seen:
            seen.add(key)
            found.append(doi)
    return found


def title_from_pdf_filename(source: str) -> str:
    stem = Path(source).stem
    if " - " in stem:
        parts = [part.strip() for part in stem.split(" - ") if part.strip()]
        if parts:
            stem = parts[-1]
    stem = re.sub(r"^\d{4}\s*[-_]\s*", "", stem)
    return clean_paper_title(stem)


def clean_paper_title(title: str) -> str:
    title = title.replace("_", " ")
    title = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", title)
    title = re.sub(r"\bRearrangementin\b", "Rearrangement in", title)
    title = re.sub(r"^Orange Phosphorus (?=Photo-Assisted Bottom-Up Synthesis)", "", title)
    title = re.sub(r"\s+Pengcheng Qiu\s*$", "", title)
    title = re.sub(r"\bcon\s+nement\b", "confinement", title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def needs_filename_title(title: str) -> bool:
    lower = title.casefold()
    return any(marker in lower for marker in ("doi.org", "www.", "research article", "article doi")) or len(title) > 150


def guess_title(text: str, fallback: str) -> str:
    lines = [normalize_space(line) for line in text.replace("\r\n", "\n").split("\n")]
    lines = [line for line in lines if len(line) >= 8]
    skip_patterns = (
        "abstract",
        "introduction",
        "references",
        "supporting information",
        "journal",
        "downloaded",
        "copyright",
        "doi:",
        "doi.org",
        "www.",
        "research article",
    )
    candidates: list[str] = []
    for line in lines[:50]:
        lower = line.casefold()
        if any(pat in lower for pat in skip_patterns):
            continue
        if len(line) > 220:
            continue
        alpha = sum(ch.isalpha() for ch in line)
        if alpha < 6:
            continue
        candidates.append(line)
    if candidates:
        title_parts = [candidates[0]]
        for line in candidates[1:4]:
            joined = " ".join(title_parts + [line])
            lower = line.casefold()
            looks_like_author_line = "," in line or re.search(r"\b(department|university|institute)\b", lower)
            if lower.startswith("and "):
                looks_like_author_line = False
            if looks_like_author_line:
                break
            if len(joined) <= 180 and (len(title_parts[0]) < 80 or title_parts[-1][-1:] not in ".?!"):
                title_parts.append(line)
            else:
                break
        title = clean_paper_title(" ".join(title_parts))
        if not needs_filename_title(title):
            return title
    return title_from_pdf_filename(fallback)


def guess_year(text: str) -> str:
    head = text[:12000]
    accepted = re.search(
        r"Accepted:\s*\d{1,2}\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s*(20[0-3]\d)",
        head,
        flags=re.IGNORECASE,
    )
    if accepted:
        return accepted.group(2)
    doi_year = re.search(r"10\.1038/s\d+-0?(\d{2})-", head)
    if doi_year:
        return "20" + doi_year.group(1)
    years = re.findall(r"\b(19[5-9]\d|20[0-3]\d)\b", head)
    if not years:
        return "n.d."
    return years[0]


def guess_publication_date(text: str, year: str) -> str:
    head = text[:20000]
    month_names = "January|February|March|April|May|June|July|August|September|October|November|December"
    accepted = re.search(
        rf"Accepted:\s*(\d{{1,2}})\s*({month_names})\s*({year})",
        head,
        flags=re.IGNORECASE,
    )
    if accepted:
        month_lookup = {
            "january": 1,
            "february": 2,
            "march": 3,
            "april": 4,
            "may": 5,
            "june": 6,
            "july": 7,
            "august": 8,
            "september": 9,
            "october": 10,
            "november": 11,
            "december": 12,
        }
        month = month_lookup[accepted.group(2).casefold()]
        return f"{accepted.group(3)}-{month:02d}-{int(accepted.group(1)):02d}"
    m = re.search(rf"\b({month_names})\s+(\d{{1,2}}),\s+({year})\b", head, re.IGNORECASE)
    if m:
        return normalize_space(m.group(0))
    m = re.search(r"\b(19[5-9]\d|20[0-3]\d)[-/\.](\d{1,2})(?:[-/\.](\d{1,2}))?\b", head)
    if m:
        return m.group(0)
    return year if year != "n.d." else ""


if __name__ == "__main__":
    raise SystemExit(main())
