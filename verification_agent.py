from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path

from agent_onenote_logger import write_agent_log


WORKSPACE_DIR = Path(__file__).resolve().parent


@dataclass
class VerificationResult:
    status: str
    checks: list[str]
    warnings: list[str]
    failures: list[str]
    completed: list[str]
    problems: list[str]
    repair_instructions: list[str]

    def to_text(self) -> str:
        lines = [f"検証ステータス: {self.status}", "", "完了したこと:"]
        lines.extend(f"- {item}" for item in self.completed) if self.completed else lines.append("- なし")
        lines.append("")
        lines.append("確認OK:")
        lines.extend(f"- {item}" for item in self.checks) if self.checks else lines.append("- なし")
        lines.append("")
        lines.append("警告:")
        lines.extend(f"- {item}" for item in self.warnings) if self.warnings else lines.append("- なし")
        lines.append("")
        lines.append("失敗/要対応:")
        lines.extend(f"- {item}" for item in self.failures) if self.failures else lines.append("- なし")
        lines.append("")
        lines.append("問題点:")
        lines.extend(f"- {item}" for item in self.problems) if self.problems else lines.append("- なし")
        lines.append("")
        lines.append("担当Agentへの修正指示:")
        lines.extend(f"- {item}" for item in self.repair_instructions) if self.repair_instructions else lines.append("- なし")
        return "\n".join(lines)


def parse_first_int(label: str, output: str) -> int | None:
    match = re.search(re.escape(label) + r"\s*[:：]\s*(-?\d+)", output)
    return int(match.group(1)) if match else None


# 検証レポート自身の否定文（例:「出力に明確なTraceback/Exceptionは見つかりません。」）に
# 反応しないよう、実際のエラー出力の形だけにマッチさせる。
ERROR_TEXT_PATTERNS = [
    re.compile(r"Traceback \(most recent call last\)"),
    re.compile(r"^\s*(?:[A-Za-z_][\w.]*)?(?:Error|Exception)\s*[:：]", re.MULTILINE),
    re.compile(r"Exception calling "),
    re.compile(r"ParserError"),
    re.compile(r"Notebook not found"),
    re.compile(r"OneNote section not found"),
    re.compile(r"失敗[:：]"),
]


def has_error_text(output: str) -> bool:
    return any(pattern.search(output) for pattern in ERROR_TEXT_PATTERNS)


def normalize_item(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip(" ・-*□☐\t")).casefold()


def extract_markdown_list_section(output: str, heading_pattern: str) -> list[str]:
    match = re.search(
        heading_pattern + r"\s*\n(?P<body>.*?)(?=\n## \d+\.|\n---\s*\n\s*## |\Z)",
        output,
        flags=re.DOTALL,
    )
    if not match:
        return []
    items = []
    for line in match.group("body").splitlines():
        stripped = line.strip()
        if stripped.startswith("- "):
            items.append(stripped[2:].strip())
    return items


def extract_raw_todo_items(output: str) -> list[str]:
    raw = output.split("## 生データ", 1)[1] if "## 生データ" in output else output
    items: list[str] = []
    in_todo = False
    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped:
            if in_todo:
                in_todo = False
            continue
        if re.fullmatch(r"[-ーｰ－―\s]+", stripped) or re.match(r"^---(?:\s+.+\s+---)?$", stripped):
            in_todo = False
            continue
        at_match = re.match(r"^[@＠]+([^\s@＠]+)(?:\s+(.*))?$", stripped, flags=re.IGNORECASE)
        if at_match:
            name = at_match.group(1).casefold().strip()
            first_arg = (at_match.group(2) or "").strip()
            in_todo = name == "todo"
            if in_todo and first_arg:
                items.append(first_arg)
            continue
        if in_todo:
            items.append(stripped.lstrip("-*・□☐ ").strip())
    deduped: list[str] = []
    seen: set[str] = set()
    for item in items:
        key = normalize_item(item)
        if key and key not in seen:
            seen.add(key)
            deduped.append(item)
    return deduped


def read_manual_todo_items(path: Path) -> list[str]:
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    match = re.search(r"## コピー用\s*\n(?P<body>.*?)(?=\n## |\Z)", text, flags=re.DOTALL)
    if not match:
        return []
    items = []
    for line in match.group("body").splitlines():
        stripped = line.strip()
        if stripped:
            items.append(stripped.lstrip("-*・□☐✅ ").strip())
    return items


def verify_common(exit_code: int | None, output: str) -> tuple[list[str], list[str], list[str]]:
    checks: list[str] = []
    warnings: list[str] = []
    failures: list[str] = []

    if exit_code is None:
        warnings.append("実行されていないため、終了コードは未確認です。")
    elif exit_code == 0:
        checks.append("終了コードが0です。")
    else:
        failures.append(f"終了コードが0ではありません: {exit_code}")

    if output.strip():
        checks.append("実行出力があります。")
    else:
        warnings.append("実行出力が空です。")

    if has_error_text(output):
        warnings.append("出力にエラー/失敗を示す語が含まれます。内容確認が必要です。")
    else:
        checks.append("出力に明確なTraceback/Exceptionは見つかりません。")

    return checks, warnings, failures


def verify_classification(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    target_pages = parse_first_int("対象ページ", output)
    failed_count = parse_first_int("失敗", output)
    existing_cards = parse_first_int("既存分類カード", output)

    if target_pages is not None:
        checks.append(f"分類対象ページ数を確認しました: {target_pages}")
    else:
        warnings.append("分類対象ページ数を読み取れませんでした。")

    if failed_count == 0:
        checks.append("分類失敗件数は0です。")
    elif failed_count is None:
        warnings.append("分類失敗件数を読み取れませんでした。")
    else:
        failures.append(f"分類失敗件数が0ではありません: {failed_count}")

    if existing_cards is not None and target_pages is not None:
        if "DRY-RUN" in output and existing_cards == target_pages:
            checks.append("dry-run上、既存分類カード数と対象ページ数が一致しています。")
        elif "DRY-RUN" in output:
            warnings.append(f"dry-run上、既存分類カード数({existing_cards})と対象ページ数({target_pages})が異なります。")


def verify_journal(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    commands_only = "[ok] コマンドだけ実行しました。" in output
    local_summary = "--local-summary" in output or "外部APIを使わずローカル抽出" in output
    saved_match = re.search(r"保存しました:\s*(.+)", output)
    if saved_match:
        path = Path(saved_match.group(1).strip())
        if path.exists():
            checks.append(f"Markdown日誌ファイルが存在します: {path}")
        else:
            failures.append(f"保存ログはありますが、Markdown日誌ファイルが見つかりません: {path}")
    elif commands_only:
        checks.append("commands-only実行のため、Markdown日誌ファイル作成は不要です。")
    else:
        warnings.append("Markdown保存先を出力から読み取れませんでした。")

    manual_todo_match = re.search(r"Google Tasks手動投入用:\s*(.+)", output)
    manual_todo_items: list[str] = []
    if manual_todo_match:
        manual_todo_path = Path(manual_todo_match.group(1).strip())
        if manual_todo_path.exists():
            checks.append(f"Google Tasks手動投入用TodoListが存在します: {manual_todo_path}")
            manual_text = manual_todo_path.read_text(encoding="utf-8-sig", errors="replace")
            for heading in ("## 分類別コピー用", "## 時刻あり", "## 表記ゆれ確認"):
                if heading in manual_text:
                    checks.append(f"Google Tasks手動投入用TodoListに {heading} セクションがあります。")
                else:
                    failures.append(f"Google Tasks手動投入用TodoListに {heading} セクションがありません。")
            manual_todo_items = read_manual_todo_items(manual_todo_path)
            if manual_todo_items:
                checks.append(f"Google Tasks手動投入用TodoListに{len(manual_todo_items)}件あります。")
            else:
                warnings.append("Google Tasks手動投入用TodoListのコピー用欄から項目を読み取れません。")
        else:
            failures.append(f"Google Tasks手動投入用TodoListが見つかりません: {manual_todo_path}")

    if "## 生データ" in output:
        checks.append("日誌末尾の生データセクションを確認しました。")
        if "Journal generated at:" in output and "updated:" in output:
            checks.append("生データ欄に日誌生成時刻とrawtext更新日時が含まれることを確認しました。")
        else:
            warnings.append("生データ欄に日誌生成時刻またはrawtext更新日時が見つかりません。")

        raw_todos = extract_raw_todo_items(output)
        summary_todos = extract_markdown_list_section(output, r"## 3\. TODO候補")
        if raw_todos:
            summary_keys = {normalize_item(item) for item in summary_todos}
            missing_todos = [item for item in raw_todos if normalize_item(item) not in summary_keys]
            if missing_todos:
                failures.append(
                    "@todoブロックの項目が日誌TODO候補に反映されていません: "
                    + " / ".join(missing_todos[:10])
                )
            else:
                checks.append(f"@todoブロック全{len(raw_todos)}件が日誌TODO候補に反映されています。")
            if manual_todo_items:
                manual_keys = {normalize_item(item) for item in manual_todo_items}
                missing_manual_todos = [item for item in raw_todos if normalize_item(item) not in manual_keys]
                if missing_manual_todos:
                    failures.append(
                        "@todoブロックの項目がGoogle Tasks手動投入用TodoListに反映されていません: "
                        + " / ".join(missing_manual_todos[:10])
                    )
                else:
                    checks.append(f"@todoブロック全{len(raw_todos)}件がGoogle Tasks手動投入用TodoListに反映されています。")
    elif not commands_only:
        warnings.append("日誌末尾の生データセクションが出力から確認できません。")

    if local_summary:
        checks.append("ローカル日誌モードの出力を確認しました。")
    elif "OneNote" in output or "onenote" in output.casefold():
        checks.append("OneNote関連の出力が含まれます。")
    else:
        warnings.append("OneNote更新確認の出力が見つかりません。別途OneNote読み取り確認が必要です。")

    if "Google Tasks" in output:
        checks.append("Google Tasks同期結果が出力に含まれます。")
        missing_match = re.search(r"Google Tasks検証:\s*確認済み(\d+)件\s*/\s*未確認(\d+)件", output)
        if missing_match:
            verified = int(missing_match.group(1))
            missing = int(missing_match.group(2))
            if missing == 0:
                checks.append(f"Google Tasks実体検証は未確認0件です: 確認済み{verified}件")
            else:
                failures.append(f"Google Tasks実体検証で未確認タスクがあります: {missing}件")
        elif sync_match := re.search(r"Google Tasks:\s*新規(\d+)件\s*/\s*重複(\d+)件\s*/\s*未同期(\d+)件", output):
            pending = int(sync_match.group(3))
            if pending == 0:
                checks.append("Google Tasksの未同期0件を確認しました。")
            else:
                failures.append(f"Google Tasksに未同期タスクがあります: {pending}件")
        elif "未同期0件" in output:
            checks.append("Google Tasksの未同期0件を確認しました。")
        else:
            warnings.append("Google Tasks同期結果がありますが、未同期/実体検証の状態を読み取れません。")


def verify_onenote_search(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    if "onenote:" in output:
        checks.append("OneNoteリンクが出力に含まれます。")
    else:
        warnings.append("OneNoteリンクが出力に見つかりません。")

    if "Page" in output or "ページ" in output:
        checks.append("検索結果らしいページ情報が含まれます。")
    else:
        warnings.append("ページ情報を出力から確認できません。")


def has_doi(output: str) -> bool:
    doi_pattern = r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b"
    return re.search(doi_pattern, output, flags=re.IGNORECASE) is not None or "doi.org/" in output.casefold()


def has_source_marker(output: str) -> bool:
    source_terms = [
        "DOI",
        "doi.org",
        "Title",
        "タイトル",
        "著者",
        "Authors",
        "Journal",
        "掲載誌",
        "Reference",
        "参考文献",
        "Source",
        "ソース",
        "Nature",
        "ACS",
        "Wiley",
        "RSC",
        "Science",
        "PubMed",
        "arXiv",
    ]
    return any(term in output for term in source_terms)


def verify_sources_required(agent: str, output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    if has_source_marker(output):
        checks.append("ソース文献を示す語句が出力に含まれます。")
    else:
        failures.append("ソース文献を示す情報が出力に見つかりません。担当Agentは再検索し、根拠ソースを提示してください。")

    if has_doi(output):
        checks.append("DOI形式の根拠文献情報が出力に含まれます。")
    elif "DOI未発見" in output or "doi未発見" in output.casefold() or "DOI not found" in output:
        failures.append("DOI未発見と出力されています。担当Agentは探索範囲を広げ、DOIまたは一次ソースを再提示してください。")
    else:
        failures.append(f"{agent} はソース文献のDOI添付が必須ですが、DOIが出力に見つかりません。担当Agentは再検索してDOIを提示してください。")


def verify_paper_search(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    verify_sources_required("論文検索Agent", output, checks, warnings, failures)
    if "onenote:" in output or "http://" in output or "https://" in output:
        checks.append("OneNoteリンクまたは外部文献リンクが出力に含まれます。")
    else:
        failures.append("OneNoteリンクまたは外部文献リンクが出力に見つかりません。")


def verify_organic_synthesis(request: str, output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    if any(term in output for term in ["Supplementary", "Supporting Information", "ESI", "SI", "特許", "PDF"]):
        checks.append("SI/ESI/PDF/特許まで確認する方針または結果が出力に含まれます。")
    else:
        warnings.append("SI/ESI/PDF/特許に関する確認が出力に見つかりません。")

    verify_sources_required("有機合成Agent", output, checks, warnings, failures)

    condition_terms = ["試薬", "溶媒", "温度", "時間", "収率", "NMR", "mmol", "当量"]
    found = [term for term in condition_terms if term in output]
    if found:
        checks.append("合成条件に関する語を確認しました: " + ", ".join(found))
    else:
        warnings.append("試薬量・溶媒・温度・収率などの条件情報が出力に見つかりません。")


def verify_experiment_note(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    required_markers = ["PPTX:", "Date:", "Pages:", "SlideCount:"]
    found = [marker for marker in required_markers if marker in output]
    if len(found) >= 3:
        checks.append("実験PPT転記の主要出力を確認しました: " + ", ".join(found))
    else:
        failures.append("実験PPT転記の主要出力(PPTX/Date/Pages/SlideCount)が不足しています。")

    if "Images:" in output:
        checks.append("OneNote画像の抽出件数が出力されています。")
    else:
        warnings.append("Images行が見つかりません。画像がある実験ページでは抽出件数を出力してください。")

    if "Snapshot:" in output or "State:" in output:
        checks.append("スナップショットまたは重複防止stateの出力を確認しました。")
    else:
        warnings.append("Snapshot/Stateの出力が見つかりません。後で差分検証しやすいよう保存先を明示してください。")


def verify_google_tasks(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    if "Google Tasks" in output or "google_todo" in output or "queue" in output.casefold():
        checks.append("Google Tasks同期またはキュー処理に関する出力を確認しました。")
    else:
        warnings.append("Google Tasks同期結果またはキュー状態の出力が見つかりません。")


def verify_device_monitor(output: str, checks: list[str], warnings: list[str], failures: list[str]) -> None:
    if "命令したLog" in output and ("Desktop" in output or "Lenovo" in output):
        checks.append("命令したLogと端末ラベルに関する出力を確認しました。")
    else:
        failures.append("命令したLog同期またはDesktop/Lenovo端末ラベルの確認出力が不足しています。")

    if "migration_check.py" in output or "再現性チェック" in output or "Migration Check" in output or "Migration check" in output:
        checks.append("再現性チェックの実行結果を確認しました。")
    else:
        failures.append("migration_check.py または再現性チェックの出力が見つかりません。")

    if "OneNote append:" in output or "命令したLogへの日次検証記録" in output:
        checks.append("監視結果を命令したLogへ記録する処理を確認しました。")
    else:
        warnings.append("監視結果の命令したLog記録確認が出力に見つかりません。")


def verify_agent_output(agent: str, request: str, output: str, exit_code: int | None) -> VerificationResult:
    checks, warnings, failures = verify_common(exit_code, output)

    if agent == "OneNote分類Agent":
        verify_classification(output, checks, warnings, failures)
    elif agent == "日誌Agent":
        verify_journal(output, checks, warnings, failures)
    elif agent == "OneNote検索Agent":
        verify_onenote_search(output, checks, warnings, failures)
    elif agent == "論文検索Agent":
        verify_paper_search(output, checks, warnings, failures)
    elif agent == "有機合成Agent":
        verify_onenote_search(output, checks, warnings, failures)
        verify_organic_synthesis(request, output, checks, warnings, failures)
    elif agent == "実験ノートAgent":
        verify_experiment_note(output, checks, warnings, failures)
    elif agent == "GoogleTasksAgent":
        verify_google_tasks(output, checks, warnings, failures)
    elif agent == "端末相互監視Agent":
        verify_device_monitor(output, checks, warnings, failures)

    if failures:
        status = "failed"
    elif warnings:
        status = "warning"
    else:
        status = "ok"

    completed = build_completed_summary(agent, exit_code, checks)
    problems = [*failures, *warnings]
    repair_instructions = build_repair_instructions(agent, request, warnings, failures)
    return VerificationResult(
        status=status,
        checks=checks,
        warnings=warnings,
        failures=failures,
        completed=completed,
        problems=problems,
        repair_instructions=repair_instructions,
    )


def build_completed_summary(agent: str, exit_code: int | None, checks: list[str]) -> list[str]:
    completed: list[str] = []
    if exit_code == 0:
        completed.append(f"{agent} は終了コード0で完了しました。")
    elif exit_code is None:
        completed.append(f"{agent} は計画確認のみで、実行はされていません。")
    else:
        completed.append(f"{agent} は実行されましたが、終了コードが0ではありません。")

    completed.extend(checks[:5])
    return completed


def build_repair_instructions(
    agent: str,
    request: str,
    warnings: list[str],
    failures: list[str],
) -> list[str]:
    if not warnings and not failures:
        return []

    instructions = ["検証Agentの警告/失敗を確認し、不足した成果物を補ってから再実行してください。"]

    if agent == "OneNote分類Agent":
        instructions.extend(
            [
                "対象ページ数、分類カード数、失敗件数を出力に明示してください。",
                "分類カード不足または失敗がある場合は全件再分類または不足分の再分類を行ってください。",
                "再実行後、分類カード数と対象ページ数の整合を確認してください。",
            ]
        )
    elif agent == "日誌Agent":
        instructions.extend(
            [
                "Markdown保存先を出力に明示してください。",
                "OneNoteページを作成/更新した後、OneNote本文を読み戻して更新確認結果を出力してください。",
                "外部送信が必要な場合は、許可済み範囲でのみ再実行してください。",
            ]
        )
    elif agent == "OneNote検索Agent":
        instructions.extend(
            [
                "検索語を広げ、ページ名、セクション、クリック可能なOneNoteリンクを必ず出力してください。",
                "結果が0件の場合は類義語・英語表記・関連語を追加して再検索してください。",
            ]
        )
    elif agent == "論文検索Agent":
        instructions.extend(
            [
                "ソースが示されていないため、再検索して根拠文献を提示してください。",
                "候補ごとにタイトル、著者、掲載誌、年、URL、DOIを明示してください。",
                "DOIが見つからない候補は採用せず、必要ならCrossref、出版社ページ、PubMed、Google Scholar、論文PDFまで検索範囲を広げてください。",
                "DOI未発見の場合は、未発見のまま終えず、探索範囲と代替一次ソースを明示して再検索してください。",
            ]
        )
    elif agent == "有機合成Agent":
        instructions.extend(
            [
                "ソースが示されていない場合は、再検索して合成条件の根拠文献とDOIを提示してください。",
                "本文だけでなくSupplementary Information、Supporting Information、SI、ESI、PDF本文、特許まで検索してください。",
                "ソース文献のDOIを必ず添付し、DOI、試薬量、mmol、当量、溶媒、温度、時間、後処理、精製、収率、NMRなどを分けて出力してください。",
                "直接条件が見つからない場合も、どこまで探したかと未解決点を明示してください。",
            ]
        )
    elif agent == "会話ログAgent":
        instructions.extend(
            [
                "Markdown、JSON、JSONLの保存先を出力してください。",
                "ユーザー発言とアシスタント返答を区別して保存してください。",
            ]
        )
    elif agent == "実験ノートAgent":
        instructions.extend(
            [
                "OneNote 2026実験/実験の対象日ページを再読込し、PPTX、Date、Pages、Images、SlideCount、Snapshot、Stateを出力してください。",
                "画像があるページではCallbackID経由の画像抽出件数も確認してください。",
                "重複防止でスキップされた場合は -Force で再実行してください。",
            ]
        )
    elif agent == "GoogleTasksAgent":
        instructions.extend(
            [
                "Google Tasksキュー件数、同期済み件数、未同期件数を出力してください。",
                "未同期が残る場合は認証、token、Google API応答、対象タスクリストを確認してください。",
            ]
        )
    elif agent == "端末相互監視Agent":
        instructions.extend(
            [
                "OpenAI_Agent1/命令したLogを再同期し、Desktop/Lenovoラベルのある最新ページを確認してください。",
                "migration_check.py --deep --write-report を再実行し、失敗項目を具体的に出力してください。",
                "監視結果を命令したLogへ追記してください。",
            ]
        )
    else:
        instructions.append(f"依頼内容を再確認し、{agent} の成果物と検証項目を出力に明示してください。")

    if "DOI" in request or "doi" in request.casefold():
        instructions.append("DOIが要求されているため、DOIまたはDOI未発見の根拠を明示してください。")

    return instructions


def verify_and_log(
    *,
    agent: str,
    request: str,
    command: list[str],
    output: str,
    exit_code: int | None,
) -> VerificationResult:
    result = verify_agent_output(agent, request, output, exit_code)
    write_agent_log(
        agent="検証Agent",
        request=f"{agent} の検証: {request}",
        decision_summary=f"{agent} の実行結果を検証し、ステータスを {result.status} と判定しました。",
        command=command,
        verification=[
            "終了コードを確認する",
            "実行出力にエラーがないか確認する",
            "担当Agentごとの必須成果物を確認する",
            "警告または失敗をOneNoteに記録する",
        ],
        output=result.to_text(),
        attach_pdfs=False,
    )
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Agent実行結果を検証し、OneNoteへ検証ログを残します。")
    parser.add_argument("--agent", required=True)
    parser.add_argument("--request", required=True)
    parser.add_argument("--command-json", default="[]")
    parser.add_argument("--output", default="")
    parser.add_argument("--exit-code", type=int)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    command = json.loads(args.command_json)
    result = verify_and_log(
        agent=args.agent,
        request=args.request,
        command=command,
        output=args.output,
        exit_code=args.exit_code,
    )
    print(result.to_text())
    return 0 if result.status != "failed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
