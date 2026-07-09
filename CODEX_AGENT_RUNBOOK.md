# Codex Agent Runbook

Last checked: 2026-07-08

This folder contains the Lenovo PC agent workflow. Use this memo to reproduce the same operations from Codex.

## Working Directory

Run commands from:

```powershell
cd "C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent\OpenAI-Agent"
```

## Python Runtime

The checked-in `.venv` was created against:

```text
C:\Users\laput\AppData\Local\Programs\Python\Python312\python.exe
```

On this Codex desktop environment that interpreter may not exist, so prefer the bundled Python:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
```

Most main scripts add `.venv\Lib\site-packages` and the bundled site-packages to `sys.path`, so the bundled Python can still use the project dependencies.

## Main Operations

## Short Launcher

Use `agent.ps1` for common actions. It automatically prefers the bundled Codex Python and falls back to `.venv` or `python`.

Important: use `agent.ps1` for user-facing operations whenever possible. The launcher automatically records successful Desktop-side actions to OneNote `OpenAI_Agent1` / `命令したLog`. This prevents the common failure mode where the main task succeeds but the audit log is forgotten.

Current rule: user-facing operations should enter through the orchestrator. `agent.ps1 journal`, `commands`, `google-sync`, `paper-search`, `experiment-onenote-day`, `daily-command-log-check`, and `device-monitor` call `orchestrator_agent.py --execute --run`, then the selected Agent runs, `verification_agent.py` checks the result, the Agent council decides whether follow-up work is needed, and the action is logged.

```powershell
.\agent.ps1 rules
.\agent.ps1 task "日誌を実行して"
.\agent.ps1 journal
.\agent.ps1 journal-todos
.\agent.ps1 commands
.\agent.ps1 orchestrator --run "日誌を更新して"
.\agent.ps1 google-sync-dry
.\agent.ps1 google-sync
.\agent.ps1 experiment-note-latest
.\agent.ps1 experiment-onenote-day -Date 2026-07-08
.\agent.ps1 sync-command-log
.\agent.ps1 log-command -Device Desktop -Summary "Desktop側で実施した内容"
.\agent.ps1 migration-check
.\agent.ps1 mutual-check
.\agent.ps1 device-monitor
```

### Desktop/Lenovo mutual monitoring

Use this on both Desktop and Lenovo to keep the two environments aligned through OneNote `OpenAI_Agent1` / `命令したLog`:

```powershell
.\agent.ps1 device-monitor
```

If the user says `相互チェックお願い`, treat it as the stronger mutual-check operation:

```powershell
.\agent.ps1 mutual-check
```

Required behavior for `相互チェックお願い`:

- Route through the orchestrator as `端末相互監視Agent`.
- Sync and inspect OneNote `OpenAI_Agent1` / `命令したLog`.
- Check whether Desktop or Lenovo command-log pages describe rule, runbook, workflow, script, or setting changes.
- Apply only necessary safe local updates on the current device.
- Rerun the monitor, verification, and Agent council after any update.
- Tell the user exactly what was changed, or that no local change was needed.
- Record the final result to `命令したLog` with the current device label.

Behavior:

- Routes the request through `端末相互監視Agent`.
- Syncs OneNote `OpenAI_Agent1` / `命令したLog`.
- Reports changed pages, other-device entries, and unlabeled entries.
- Runs `migration_check.py --deep --write-report`.
- Writes the monitoring result back to `命令したLog` with the current device label.

If `Changed count` is greater than zero, read the changed page Markdown files listed in `agent_workspace\司令塔Agent\onenote_command_log\changed_pages.json`. Apply any necessary runbook or script updates on the current machine, then rerun `.\agent.ps1 device-monitor`.

### Append experiment note to PowerPoint

Append a structured experiment-note slide to `Experiment.pptx`:

```powershell
.\agent.ps1 experiment-note -Title "2026-07-08 STM sample preparation" -Sample "Sample ID ..." -Objective "..." -Procedure "..." -Conditions "..." -Observations "..." -Results "..." -NextActions "..."
```

Append from the latest journal Markdown:

```powershell
.\agent.ps1 experiment-note-latest
```

Outputs:

- `Experiment.pptx`
- `agent_workspace\実験ノートAgent\logs\*.md`

### Copy OneNote experiment day into PowerPoint

Create a date slide in `Experiment.pptx`, then copy all text and images from OneNote experiment notebook pages for the same date.

```powershell
.\agent.ps1 experiment-onenote-day -Date 2026-07-08
```

Default source:

- OneNote notebook: `2026実験`
- OneNote section: `実験`

Optional override:

```powershell
.\agent.ps1 experiment-onenote-day -Date 2026-07-08 -NotebookName "2026実験" -SectionName "有機合成"
```

Behavior:

- Finds OneNote pages in the experiment notebook containing the same date.
- Adds a date summary slide.
- Adds text continuation slides when one slide is not enough.
- Extracts OneNote embedded images and adds image slides.
- Saves a JSON snapshot and dedupe state under `agent_workspace\実験ノートAgent\onenote_to_ppt`.
- Skips duplicate content unless `-Force` is passed.

### Sync OneNote command log

Read OneNote `OpenAI_Agent1` / `命令したLog` and mirror pages locally:

```powershell
.\agent.ps1 sync-command-log
```

Outputs:

- `agent_workspace\司令塔Agent\onenote_command_log\latest_sync_report.md`
- `agent_workspace\司令塔Agent\onenote_command_log\changed_pages.json`
- `agent_workspace\司令塔Agent\onenote_command_log\pages\*.md`

When `Changed count` is greater than zero, read the changed page Markdown files and update this project or runbook as needed.

### Append Desktop command log

Record work done on this desktop into OneNote `OpenAI_Agent1` / `命令したLog`, with a local fallback copy:

```powershell
.\agent.ps1 log-command -Device Desktop -Summary "Desktop側で実施した内容" -Actions "agent.ps1を更新; OneNote同期スクリプトを追加" -Files "agent.ps1; sync_onenote_command_log.ps1"
```

Local fallback:

- `agent_workspace\司令塔Agent\command_log_outbox\*.md`

Use `Device Lenovo` when recording Lenovo-side work manually, and `Device Desktop` for this machine. Keeping the device label makes later verification much easier.

### Command Log Format

Every entry in OneNote `OpenAI_Agent1` / `命令したLog` should keep this machine-readable shape so Desktop and Lenovo can monitor each other:

```md
## Summary
## Actions
## Files
## Verification
## Required On Other Device
## Next steps
```

Rules:

- `Device` must be `Desktop` or `Lenovo`.
- `Files` should list touched or generated files relative to `OpenAI-Agent` when possible.
- `Verification` should include concrete checks such as `py_compile OK`, `migration_check OK`, `OneNote write OK`, or generated output paths.
- `Required On Other Device` should say what the other PC must do. Use `なし` only when no cross-device action is needed.
- `mutual_command_log_monitor.py` reads `Required On Other Device` and reports it under `Required On This Device`.

### Auto command logging policy

Cause of the previous miss:

- The actual work and the `命令したLog` entry were separate manual steps.
- When the work succeeded, it was easy to send the final response before running `log-command`.

Countermeasure:

- User-facing `agent.ps1` actions now go through the orchestrator where possible, then call `append_onenote_command_log.ps1` automatically after successful major actions.
- Auto-logged actions include `journal`, `journal-todos`, `commands`, `orchestrator`, `google-sync`, `google-auth`, `paper-search`, `experiment-note`, `experiment-note-latest`, and `experiment-onenote-day`.
- `daily-command-log-check` and `device-monitor` log through `端末相互監視Agent` itself. `rules`, `sync-command-log`, `log-command`, and `migration-check` do not auto-log by default because they are read-only/status or logging actions.

Operational rule:

- Prefer `.\agent.ps1 journal` instead of running `summarize_note5.py` directly.
- Prefer `.\agent.ps1 experiment-onenote-day -Date YYYY-MM-DD` instead of running the underlying PowerShell script directly.
- If a direct script run is unavoidable, immediately run `.\agent.ps1 log-command -Device Desktop -Summary "..."`

### Show Research OS command rules

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\summarize_note5.py --show-rules
```

### Generate today's journal from `rawtext`

This sends `rawtext` contents to the OpenAI API.

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\summarize_note5.py --force
```

Useful variants:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\summarize_note5.py --force --sync-google-todos
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\summarize_note5.py --commands-only --execute-commands --sync-google-todos
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\summarize_note5.py --watch --delta-only --append-output
```

Outputs are written to `日誌`.

### Route a request through the orchestrator

Planning only:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\orchestrator_agent.py "日誌を更新して"
```

Execute selected agent:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\orchestrator_agent.py --run "日誌を更新して"
```

Write-capable execution:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\orchestrator_agent.py --execute --run "日誌を更新して"
```

The orchestrator routes requests to 日誌Agent, OneNote検索Agent, 論文検索Agent, 有機合成Agent, 会話ログAgent, and 検証Agent, then records logs under `agent_workspace`.

### Sync queued Google Tasks

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\sync_google_todo_queue.py --dry-run
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\sync_google_todo_queue.py
```

Queue files:

- `日誌\google_todo_queue.jsonl`
- `日誌\google_todo_synced.jsonl`

Credentials:

- `credentials.json`
- `token_google_tasks.json`

### Refresh Google Tasks OAuth token

This may open a browser or require account interaction.

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\setup_google_tasks.py
```

### Daily paper search

This calls Crossref and OpenAI API scoring.

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\daily_paper_search.py
```

Output:

- `papers\YYYY-MM-DD_paper_list.xlsx`

### OneNote / Paper / Organic synthesis agents

These PowerShell agents use OneNote Desktop COM automation:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\onenote_search_agent.ps1 -Query "太陽電池"
powershell -ExecutionPolicy Bypass -File .\tools\paper_search_agent.ps1 -Query "TMS ethynyl triazine DOI"
powershell -ExecutionPolicy Bypass -File .\tools\organic_synthesis_agent.ps1 -Query "2-butoxynaphthalene synthesis DOI"
```

They require OneNote Desktop and COM access on Windows.

If `New-Object -ComObject OneNote.Application` fails with `0x80070520` / "A specified logon session does not exist":

- Start OneNote Desktop in the same Windows user session.
- Sign in and open/sync `OpenAI_Agent1` before rerunning `.\agent.ps1 device-monitor`.
- If the command is running from a background/non-interactive automation session, treat OneNote write/sync as unavailable for that run and check `agent_workspace\司令塔Agent\command_log_outbox\*.md` for the local fallback entry.
- Rerun `.\agent.ps1 migration-check` or `.\agent.ps1 device-monitor` after OneNote Desktop is visible and synced.

## Agent Workspaces

Agent logs and records are under:

- `agent_workspace\司令塔Agent`
- `agent_workspace\日誌Agent`
- `agent_workspace\OneNote検索Agent`
- `agent_workspace\有機合成Agent`
- `agent_workspace\検証Agent`
- `agent_workspace\会話ログAgent`

## Environment / Auth

`.env` should contain:

```text
OPENAI_API_KEY=...
GOOGLE_TODO_SYNC=1
GOOGLE_TASKS_CREDENTIALS=credentials.json
GOOGLE_TASKS_TOKEN=token_google_tasks.json
GOOGLE_TASKS_LIST_ID=@default
```

Do not paste the real `.env`, OAuth client secret, or token contents into chat.

## Current Migration Notes

- `.venv\Scripts\python.exe` is not reliable on this machine because `pyvenv.cfg` points to a missing Lenovo-side Python path.
- `sync_google_todo_queue.py` and `setup_google_tasks.py` were updated to add the project and bundled site-packages path before importing dependencies.
- `setup_google_tasks.py` was updated to import `summarize_note5` instead of the old `summarize_note4`.
- Some OneDrive files are cloud placeholders and may require elevated/unsandboxed access or local hydration before Python can import them normally.
- OneNote automation primarily uses PowerShell COM. `migration_check.py --deep` verifies `OneNote.Application` through PowerShell COM, so bundled Python does not need `win32com` for the normal workflow.

## Health Check

Run:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\migration_check.py
```

Deep check, including OneNote PowerShell COM:

```powershell
C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\migration_check.py --deep --write-report
```

Expected as of 2026-07-08:

- Most project files and auth files are present.
- `GOOGLE_TODO_SYNC=1` is set in `.env`, so @todo entries are pushed to Google Tasks by default when command execution runs.
- Deep migration check should report `WARN 0 / FAIL 0`; OS information may remain as `INFO`.
- OneNote COM is checked through PowerShell COM, not Python `win32com`, for the normal workflow.
- Possible failure if OneDrive placeholder files cannot be imported without hydration/elevated access.
