# Claude Agent Runbook

Last checked: 2026-07-09

This folder is the Claude-based port of the Codex `OpenAI-Agent` research-automation system
(`..\..\Codex\OpenAI-Agent`). Use this memo to reproduce the same operations from Claude Code.

## 課金ポリシー（重要・ユーザー方針 2026-07-09）

ユーザー（小島先生）はAPI課金に慎重で、課金ポイントを自分で管理したい意向。**Claude API課金と
外部送信は既定でオフ**にしてある。課金が発生する／発生し得る変更を加える前には、必ずユーザーに
明示して判断を仰ぐこと。

現在の既定（すべて課金なし・外部送信なし）:

| 機能 | 既定 | 課金オプトインの方法 |
|---|---|---|
| 日誌の要約（`agent.ps1 journal` 等） | **ローカル抽出**（Claude不使用） | `summarize_note5.py` を `--local-summary` 無しで直接実行 |
| `@ask`（Claude＋web検索） | **無効**（`--skip-ask`） | journal/commandルートから `--skip-ask` を外す |
| Google Tasks送信 | **停止**（手動投入用リストのみ） | `.env` で `GOOGLE_TODO_SYNC=1`、または `sync_google_todo_queue.py --force-sync` |
| 論文の関連度スコアリング（`daily_paper_search.py`） | **無効** | `.env` で `PAPER_CLAUDE_SCORING=1`、または `--use-claude` |
| 論文PDF要約（`summarize_paper.py`） | **無効** | `.env` で `PAPER_CLAUDE_SUMMARY=1` |

課金なしで使える主なコマンド: `journal` / `journal-local` / `commands`（@todoは手動リスト化）、
`migration-check`、`device-monitor`、`sync-command-log`。
Crossref（論文メタデータ取得）は無料。ルーティング・検証・OneNote記録はすべてローカル処理。

## Working Directory

Run commands from:

```powershell
cd "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\Claude\OpenAI-Agent"
```

## Python Runtime

This project uses its own dedicated `.venv`, created against:

```text
C:\Users\laput\AppData\Local\Programs\Python\Python312\python.exe
```

Setup (one time per machine):

```powershell
C:\Users\laput\AppData\Local\Programs\Python\Python312\python.exe -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -r requirements-google-tasks.txt
```

Unlike the Codex version, there is no bundled-runtime fallback — `agent.ps1` uses the project
`.venv` directly, falling back to whatever `python` is on `PATH` only if `.venv` doesn't exist yet.

## OneNote destinations

Use the following destinations for new journal and paper-summary pages (2026-07-16、Codex側の移行に追随):

- Journal: `FarEasternTribe` / `日誌`
- Paper summaries: `FarEasternTribe` / `paper_summarize`
- Research paper summaries formerly under `2026実験`: `FarEasternTribe` / `論文要約`
- Experiment pages formerly under `2026実験`: `FarEasternTribe` / `実験`

`OpenAI_Agent1` remains the destination for command and Agent audit logs (`命令したLog_Claude`);
do not write new journal or paper-summary pages there. The notebook can be overridden with the
`ONENOTE_NOTEBOOK_NAME` environment variable (`summarize_note5.py`), `-NotebookName`
(`append_onenote_experiment_day_to_ppt.ps1`, `experiment_tracker.ps1`), etc.
Note: `tools/add_synthesis_pdf_to_onenote.py`（有機合成）はCodex版と同じく既定 `2026実験` のまま
（有機合成セクションは未移行）。`experiment_tracker.ps1`（実験ボード）は実験ページ移行に合わせ
`FarEasternTribe` 既定に変更済み。

## Short Launcher

Use `agent.ps1` for common actions, same action names as the Codex version:

```powershell
.\agent.ps1 rules
.\agent.ps1 task "日誌を実行して"
.\agent.ps1 journal
.\agent.ps1 journal-local
.\agent.ps1 journal-todos
.\agent.ps1 commands
.\agent.ps1 orchestrator --run "日誌を更新して"
.\agent.ps1 google-sync-dry
.\agent.ps1 google-sync
.\agent.ps1 google-auth
.\agent.ps1 paper-search
.\agent.ps1 experiment-note-latest
.\agent.ps1 experiment-onenote-day -Date 2026-07-09
.\agent.ps1 journal-task-status
.\agent.ps1 sync-command-log
.\agent.ps1 log-command -Summary "この端末で実施した内容"   # Desktop/Lenovoは自動判定
.\agent.ps1 migration-check
.\agent.ps1 status
.\agent.ps1 secrets-status
.\agent.ps1 backup
.\agent.ps1 restore-check .\backups\agent-config-YYYYMMDD-HHMMSS.zip
.\agent.ps1 test
.\agent.ps1 mutual-check
.\agent.ps1 device-monitor
```

User-facing operations route through `orchestrator_agent.py --execute --run`, which runs the
selected Agent, calls `verification_agent.py`, holds an "Agent council", and logs the result to
OneNote (`OpenAI_Agent1` / `命令したLog_Claude`) automatically. Prefer `agent.ps1` over invoking
scripts directly so this logging isn't missed.

### Desktop/Lenovo mutual monitoring

```powershell
.\agent.ps1 device-monitor
```

Stronger check when the user says `相互チェックお願い`:

```powershell
.\agent.ps1 mutual-check
```

Same behavior as the Codex version (syncs OneNote, diffs Desktop/Lenovo entries, runs
`migration_check.py --deep --write-report`, writes the result back), except it reads/writes the
**`命令したLog_Claude`** section of the `OpenAI_Agent1` notebook, not `命令したLog`. This keeps
Claude-origin entries separate from the Codex version's entries — the two tools do not currently
read each other's logs. `append_onenote_command_log.ps1` tags entries with `Actor: Claude`.

**Known-benign warning**: `verify_device_monitor` in `verification_agent.py` warns when
`"OneNote append:"` isn't found in the monitor's stdout — this happens whenever OneNote Desktop
isn't running/signed-in in an interactive session (headless/background runs). This is expected in
that situation, not a bug: `build_repair_command()` intentionally returns `None` for
`端末相互監視Agent` (unlike the Codex version, which used to return the same command unchanged
and re-run it pointlessly — see "Fixed from the Codex version" below), so the orchestrator reports
the warning once instead of looping. To actually clear the warning, start OneNote Desktop, sign
in, and open/sync `OpenAI_Agent1`, then re-run `device-monitor`.

### Append experiment note to PowerPoint

```powershell
.\agent.ps1 experiment-note -Title "..." -Sample "..." -Objective "..." -Procedure "..." -Conditions "..." -Observations "..." -Results "..." -NextActions "..."
.\agent.ps1 experiment-note-latest
.\agent.ps1 experiment-onenote-day -Date 2026-07-09
```

Behavior is unchanged from the Codex version (see `CODEX_AGENT_RUNBOOK.md` in the Codex folder for
full details) — outputs go to `Experiment.pptx` and `agent_workspace\実験ノートAgent\`.

### Sync OneNote command log

```powershell
.\agent.ps1 sync-command-log
```

Mirrors `OpenAI_Agent1` / `命令したLog_Claude` pages to
`agent_workspace\司令塔Agent\onenote_command_log\`.

### Command Log Format

Same structure as the Codex version:

```md
## Summary
## Actions
## Files
## Verification
## Required On Other Device
## Next steps
```

`Device` must be `Desktop` or `Lenovo`. `Files` relative to `OpenAI-Agent`. `Required On Other
Device` read by `mutual_command_log_monitor.py` on the other machine.

### Show Research OS command rules

```powershell
.\.venv\Scripts\python.exe .\summarize_note5.py --show-rules
```

### Source memo folder (`rawtext`)

When run through `agent.ps1 journal` / the orchestrator, the journal and `@command` agents read
their source memos from the **shared rawtext folder in the old OpenAI_Agent tree**, not this
project's local `rawtext\`:

```text
C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent\OpenAI-Agent\rawtext
```

This is set as `RAWTEXT_DIR` (default `DEFAULT_RAWTEXT_DIR`) in `orchestrator_agent.py`; override
it with a `RAWTEXT_DIR` environment variable if the memos move. Only top-level `*.txt` files are
read (subfolders like `old\` are ignored). Since 2026-07-15, `summarize_note5.py` also defaults to
this shared folder when run directly without a path argument (`DEFAULT_INPUT = os.getenv("RAWTEXT_DIR",
DEFAULT_SHARED_RAWTEXT)`), so every journal-update path — orchestrator or direct — reads the same
shared rawtext. Pass an explicit path (or set `RAWTEXT_DIR`) only to override.

### Generate today's journal from `rawtext`

**既定はローカル要約（課金なし）。** 2026-07-16以降、`summarize_note5.py` 自体の既定が
`--local-summary` になった（直接実行でもAPIは呼ばれない）。`agent.ps1 journal` は差分実行が
既定で、`--force` は全件再実行が本当に必要なときだけ付ける。同等の直接実行:

```powershell
.\.venv\Scripts\python.exe .\summarize_note5.py "<RAWTEXT_DIR>" --no-sync-google-todos --manual-google-tasks --skip-ask
```

**Claude要約を使う場合（課金あり・明示オプトイン）** — `--api-summary` を明示して直接実行する
（`ANTHROPIC_API_KEY` が必要）。実行前にユーザーの承認を得ること:

```powershell
.\.venv\Scripts\python.exe .\summarize_note5.py "<RAWTEXT_DIR>" --api-summary
```

### OneNote To Do tags from `@Todo` blocks (2026-07-16、Codex移植)

When the source contains an `@Todo` / `＠Todo` block, those items are appended to the very end of
the OneNote journal page as real unchecked OneNote **To Do** tags (the same checkboxes applied by
`Ctrl+1`). Voice/OCR input written as `& Todo` is treated as the same marker.

When an existing journal page is updated, its current OneNote task list is authoritative:
existing tasks keep their text, order, and checked state; only newly encountered source Todos are
appended. Previously sent source keys are tracked in `日誌/onenote_todo_source_state.json`, so a
typo corrected manually in OneNote is not reverted and the old rawtext spelling is not re-added.

### Check journal task progress

Treat the actual OneNote To Do tag state as the source of truth. For questions like
`あと何が残ってる？` / `終わったタスクは？`, read the page first:

```powershell
.\agent.ps1 journal-task-status
.\agent.ps1 journal-task-status -Date 2026-07-16
```

Do not infer completion from the Markdown file because checkbox state changes later in OneNote.

### OneNote handwritten-note transcription trigger

`XXX、文字起こし` → transcribe the **latest** synchronized state of the OneNote page `XXX`.
A repeated request always captures the current latest state; never reuse an earlier PDF or transcription just because an output page already exists.

1. For a non-date title, exact-title search across all notebooks (multiple matches → stop and ask).
   For an eight-digit date `XXX`, compare the exact-date source page and every `XXX_手書き文字起こし…` page and pick the eligible one with the newest OneNote `lastModifiedTime`. An automation-created transcription page is registered under `.agent_runtime\transcription_outputs`, so an untouched output excludes itself; it becomes eligible again once the user edits it.
2. Export to PDF. Non-date titles use exact-title selection; eight-digit dates add `-LatestForDate`:
   `powershell -ExecutionPolicy Bypass -File .\tools\export_onenote_page_pdf.ps1 -PageTitle "XXX" -Out ".\tmp\pdfs\XXX_handwriting.pdf"` (append `-LatestForDate` for a date). Confirm `SyncRequested=True`, `ResolvedPageTitle`, `ResolvedPageId`, `PageLastModified`.
3. Render pages at 300 DPI, transcribe printed text and handwriting in reading order; unclear text → `［要確認］`.
   - Treat handwritten diagrams, reaction schemes, structural formulas, plots, meaningful tables, and equations as visual evidence — do not replace them with guessed prose. Crop each region from the 300 DPI render (keep labels/arrows/subscripts) into the transcription workspace.
   - Put a marker on its own line at the matching point in reading order: `[[FIGURE:relative\path.png|desc]]` for diagrams/schemes, `[[EQUATION:relative\path.png|desc]]` for equations, `[[IMAGE:relative\path.png|desc]]` for ordinary photos.
4. Save the UTF-8 transcription to `agent_workspace\OneNote検索Agent\XXX_手書き文字起こし.md`.
5. Create a new page `XXX_手書き文字起こし` with `tools\create_onenote_text_page.ps1` (never edit the source page). That wrapper delegates to the hardened shared writer `tools\create_onenote_text_image_page.ps1`, which embeds FIGURE/EQUATION/IMAGE markers as real OneNote images in source order, reads back the created page's own XML (never rebuilds a `<one:Page>` from the ID alone), retries `UpdatePageContent` up to 5×, verifies title/text-lines/image-count/`VisualOrder`, and removes an empty page it created on failure. If the output title already exists, keep it and use the next suffix.
6. Read back and verify title, non-empty line count, first/last line, image/figure/equation counts, and `VisualOrder`; confirm each visual stayed adjacent to its explanatory text; re-export once more, and if the source `PageLastModified` advanced repeat capture. Delete temp files; log to 命令したLog_Claude.

`XXX、更新分だけ文字起こし` → run `tools\detect_onenote_transcription_updates.ps1 -PageTitle "XXX"`,
transcribe only the changed-region images in `manifest.json`
(`ChangedRegionCount=0` → report no update), append to a new page
`XXX_手書き文字起こし_更新_YYYYMMDD-HHmmss`, then commit the baseline with `-CommitBaseline`
**only after** read-back verification. Baseline lives under
`.agent_runtime\transcription_state\<SHA-256 title key>\baseline`.
Note: `pdftoppm` (Poppler) is required; the Claude fork falls back to the Codex runtime's bundled
Poppler on this PC if it is not on PATH.

### Route a request through the orchestrator

```powershell
.\.venv\Scripts\python.exe .\orchestrator_agent.py "日誌を更新して"                 # plan only
.\.venv\Scripts\python.exe .\orchestrator_agent.py --run "日誌を更新して"           # execute
.\.venv\Scripts\python.exe .\orchestrator_agent.py --execute --run "日誌を更新して" # execute + writes
```

### Daily paper search

Calls Crossref and the Claude API for relevance scoring:

```powershell
.\.venv\Scripts\python.exe .\daily_paper_search.py
```

Output: `papers\YYYY-MM-DD_paper_list.xlsx` (scoring columns: `Relevance score`, `Claude reason`,
`Research connection`, `Recommended action`, `Claude keywords`).

### OneNote / Paper / Organic synthesis agents

Same as the Codex version — these PowerShell agents use OneNote Desktop COM automation and
require OneNote Desktop running, signed in, and synced in the same interactive Windows session:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\onenote_search_agent.ps1 -Query "太陽電池"
powershell -ExecutionPolicy Bypass -File .\tools\paper_search_agent.ps1 -Query "TMS ethynyl triazine DOI"
powershell -ExecutionPolicy Bypass -File .\tools\organic_synthesis_agent.ps1 -Query "2-butoxynaphthalene synthesis DOI"
```

If `New-Object -ComObject OneNote.Application` fails with `0x80070520`: start OneNote Desktop in
the same session, sign in, open/sync `OpenAI_Agent1`, then re-run.

## Environment / Auth

Secrets should live outside OneDrive and the Git workspace under
`%LOCALAPPDATA%\OpenAI-Agent\secrets` (shared with the Codex-side agent). The external
`agent.env` should contain:

```text
# Optional: only needed for explicitly requested --api-summary or other API-based tools.
ANTHROPIC_API_KEY=...
GOOGLE_TODO_SYNC=1
GOOGLE_TASKS_CREDENTIALS=%LOCALAPPDATA%\OpenAI-Agent\secrets\credentials.json
GOOGLE_TASKS_TOKEN=%LOCALAPPDATA%\OpenAI-Agent\secrets\token_google_tasks.json
GOOGLE_TASKS_LIST_ID=@default
```

Run `.\agent.ps1 secrets-status` to verify locations without displaying contents.
**Migration status (2026-07-16):** the external secret directory does not exist yet on this
Desktop; the legacy workspace `.env` / `credentials.json` / `token_google_tasks.json` are still
in place and still work (`agent_config.py` loads the external `agent.env` first, then falls back
to the workspace `.env`). Moving the actual secret files requires explicit user confirmation —
see `AGENT_OPERATION_POLICY.md`. `credentials.json` and `token_google_tasks.json` were copied
from the Codex version so both ports share the same Google Tasks account without re-running
OAuth. Do not paste `.env`/`agent.env`, the OAuth client secret, or token file contents into chat.

## Reliability, status, and recovery

Use `agent.ps1` as the single user-facing entry point. Repeat mutating actions with the same
arguments are suppressed for ten minutes (`--force` to override).

```powershell
.\agent.ps1 status
.\agent.ps1 test
.\agent.ps1 backup
.\agent.ps1 restore-check .\backups\agent-config-YYYYMMDD-HHMMSS.zip
.\agent.ps1 secrets-status
```

- `.agent_runtime\state.json` is the single local view of running, succeeded, failed, and repeated attempts.
- Shared Python and PowerShell helpers live in `agent_common.py` and `tools\agent_common.ps1`; keep
  device detection, subprocess execution, OneNote path formatting, and DOI parsing there instead of
  copying implementations.
- `backup` creates a secret-free configuration/source ZIP under `backups`.
- `restore-check` only lists and validates restore candidates. Applying a restore requires the
  lower-level command's explicit `--apply --confirm-restore` pair and first makes a pre-restore safety copy.
- Automatic versus confirmation-required operations are defined in `AGENT_OPERATION_POLICY.md`.
- Small regression tests are under `tests` and are run through `.\agent.ps1 test`.

## Differences from the Codex version

- LLM calls go through `llm_client.py` (Anthropic `messages.create`, model `claude-sonnet-5`)
  instead of `openai`'s Responses API — see `summarize_note5.py`, `daily_paper_search.py`.
- `migration_check.py` checks `ANTHROPIC_API_KEY` / the `anthropic` package instead of
  `OPENAI_API_KEY` / `openai`, and no longer has a Codex-bundled-runtime fallback check.
- No Codex bundled-runtime interpreter fallback anywhere (`agent.ps1`,
  `summarize_note5.py`, `daily_paper_search.py`, `migration_check.py`) — this project always uses
  its own `.venv`.
- **Fixed from the Codex version**: `build_repair_command()` in `orchestrator_agent.py` used to
  return the unmodified command for `端末相互監視Agent`, causing an infinite
  warn → pointless-rerun → warn loop every time OneNote COM wasn't available. It now returns
  `None` (no correctable command), matching the fallback behavior for other unhandled agents.
- Desktop/Lenovo command log lives in a separate OneNote section, `命令したLog_Claude`, so it
  doesn't mix with the Codex version's `命令したLog` entries.

## Health Check

```powershell
.\.venv\Scripts\python.exe .\migration_check.py
.\.venv\Scripts\python.exe .\migration_check.py --deep --write-report
```

Expect `WARN 0 / FAIL 0` on a healthy machine (OS info may show as `INFO`). Run after any change
to imports, paths, or credentials handling, and as part of `device-monitor`/`mutual-check`.
