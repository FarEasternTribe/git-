# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> **Provenance**: This repository is the Claude-development copy, forked on 2026-07-09 from the live
> Codex-operated system at
> `C:\Users\laput\OneDrive - Kyoto University\2-総合デスクトップ(2024)\0000000000OpenAI_Agent_Claude\Codex\OpenAI-Agent`.
> The original folder keeps running the production workflow (Desktop/Lenovo mutual monitoring, hourly logs);
> develop here, and port changes back deliberately. Runtime data directories (`rawtext/`, `日誌/`,
> `agent_workspace/`, etc.) were intentionally not copied and are gitignored — they will be recreated on first
> run, or copy sample data from the original folder when needed.

## Project overview

This is a personal research-automation agent system (Windows, PowerShell + Python) that a Kyoto University
researcher runs on two PCs ("Desktop" and "Lenovo"). A natural-language request (Japanese) is routed by
`orchestrator_agent.py` ("司令塔Agent" / control-tower agent) to one of several specialized agents that:
summarize daily notes into a journal, search/classify OneNote pages, search literature, plan organic-synthesis
routes, sync TODOs to Google Tasks, transcribe OneNote experiment notes into `Experiment.pptx`, and keep the two
machines in sync via a shared OneNote command log. There is no web/app frontend — everything is CLI scripts
invoked from PowerShell, OpenAI API calls, and OneNote Desktop COM automation.

See `CODEX_AGENT_RUNBOOK.md` for the authoritative, up-to-date operational runbook (working directory, Python
runtime selection, every `agent.ps1` action, command-log format, and known migration caveats). Prefer updating
that file over duplicating instructions here when operational details change.

## Running commands

There is no build step, package manifest, or automated test suite in this repo (no `package.json`,
`pyproject.toml`, or `tests/`). "Verification" is done by the agents themselves via `verification_agent.py` and
by running `migration_check.py`.

### Python interpreter

Scripts are written to run under either interpreter; most add the relevant `site-packages` to `sys.path`
themselves:

- Bundled Codex runtime (preferred on this machine): `C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`
- Project venv (fallback, may be stale/mismatched ABI): `.venv\Scripts\python.exe`

`agent.ps1` picks the right interpreter automatically — prefer it over invoking Python directly.

### Preferred entry point: `agent.ps1`

Always run user-facing operations through `agent.ps1` rather than the underlying scripts directly. Its actions
route through the orchestrator (`orchestrator_agent.py --execute --run`), which executes the chosen agent, runs
`verification_agent.py`, holds an "Agent council" to decide on follow-up work, and — for most actions — logs the
result to OneNote (`OpenAI_Agent1` / `命令したLog`) automatically:

```powershell
.\agent.ps1 rules                                     # show the @command spec (研究OS)
.\agent.ps1 task "日誌を実行して"                      # free-form natural-language request
.\agent.ps1 journal                                    # summarize rawtext/ into today's journal
.\agent.ps1 journal-local                              # same, but no external API (local extraction only)
.\agent.ps1 journal-todos                              # journal + Google Tasks sync
.\agent.ps1 commands                                   # execute @-commands found in rawtext
.\agent.ps1 orchestrator --run "<request>"             # route a request without agent.ps1's extra logging
.\agent.ps1 google-sync-dry / google-sync              # Google Tasks queue sync (dry-run / real)
.\agent.ps1 google-auth                                # refresh Google Tasks OAuth token
.\agent.ps1 paper-search                                # literature search agent
.\agent.ps1 experiment-note-latest                     # append latest journal as a slide to Experiment.pptx
.\agent.ps1 experiment-onenote-day -Date YYYY-MM-DD     # copy a OneNote experiment-notebook day into the PPTX
.\agent.ps1 sync-command-log                           # mirror OneNote 命令したLog pages locally
.\agent.ps1 log-command -Device Desktop -Summary "..."  # manually append a command-log entry
.\agent.ps1 migration-check                             # environment/health check
.\agent.ps1 device-monitor                              # Desktop/Lenovo mutual sync + reproducibility check
.\agent.ps1 mutual-check                                # stronger version, triggered by "相互チェックお願い"
```

### Health check ("closest thing to a test suite")

```powershell
<python> .\migration_check.py                          # quick check
<python> .\migration_check.py --deep --write-report     # also checks OneNote COM via PowerShell; writes a report
```

Expect `WARN 0 / FAIL 0` on a healthy machine; OS info may show as `INFO`. Run this after any change to imports,
paths, or credentials handling, and as part of `device-monitor`/`mutual-check`.

### Running a single agent script directly

Only do this when `agent.ps1` has no matching action, and immediately follow up with
`.\agent.ps1 log-command -Device Desktop -Summary "..."` so the cross-device audit log isn't missed (this was a
recurring failure mode — see "Auto command logging policy" in the runbook).

```powershell
<python> .\summarize_note5.py --show-rules
<python> .\summarize_note5.py --force
<python> .\orchestrator_agent.py "日誌を更新して"                 # plan only, no execution
<python> .\orchestrator_agent.py --run "日誌を更新して"           # execute
<python> .\orchestrator_agent.py --execute --run "日誌を更新して" # execute + allow writes
<python> .\sync_google_todo_queue.py --dry-run
<python> .\daily_paper_search.py
powershell -ExecutionPolicy Bypass -File .\tools\onenote_search_agent.ps1 -Query "..."
powershell -ExecutionPolicy Bypass -File .\tools\paper_search_agent.ps1 -Query "..."
powershell -ExecutionPolicy Bypass -File .\tools\organic_synthesis_agent.ps1 -Query "..."
```

The OneNote-COM-based `.ps1` agents under `tools/` require OneNote Desktop to be running, signed in, and synced
in the same interactive Windows session; they will fail with `0x80070520` otherwise (see runbook for recovery
steps).

## Architecture

### Request routing (`orchestrator_agent.py`)

`route_request()` classifies a raw Japanese request by keyword matching (no LLM call for routing itself) into a
`Route` dataclass (agent name, reason, shell command, list of verification checks the agent must satisfy,
whether it writes/sends externally), in roughly this priority order: device-monitor → conversation-log →
`@command` execution → Google Tasks sync → experiment-PPT → paper search → organic synthesis → OneNote
classification → journal/summarize → OneNote search → fallback ("司令塔Agent" with no route).

`main()` then, when `--run`/`--execute` is set: runs the routed command as a subprocess, writes the routing
decision and agent output to OneNote via `agent_onenote_logger.write_agent_log`, calls `verification_agent.py`
(mandatory for `有機合成Agent`, `論文検索Agent`, `実験ノートAgent`, `端末相互監視Agent` regardless of `--no-verify`),
auto-repairs and re-runs with broadened parameters if verification isn't `ok` (`build_repair_command`), then runs
an "Agent council" (`build_agent_council_decision`) that can trigger one more follow-up execution (e.g. if a
paper/synthesis agent's output lacks a DOI), and finally appends a structured entry to the conversation log
(`conversation_log_agent.py`) and to `tools/orchestrator_agent_log.jsonl`.

### Verification (`verification_agent.py`)

Pure text/regex-based verification per agent type (`verify_journal`, `verify_paper_search`,
`verify_organic_synthesis`, `verify_experiment_note`, `verify_device_monitor`, etc.) that inspects the agent's
stdout/stderr for required markers (e.g. DOI regex `10\.\d{4,9}/...`, `PPTX:`/`Date:`/`Pages:`/`SlideCount:`
markers, `## 生データ` raw-data section, Google Tasks sync counts) and produces a `status` of `ok` / `warning` /
`failed` plus concrete repair instructions fed back into the orchestrator's auto-repair loop.

### OneNote-based audit trail

Nearly every agent action is recorded as a page/log entry in OneNote notebook `OpenAI_Agent1`
(`agent_onenote_logger.write_agent_log`, PowerShell COM under the hood), with a local Markdown fallback under
`agent_workspace/<Agent名>/logs/`. The special section `命令したLog` (`COMMAND_LOG_SECTION`) is the cross-device
sync point: each entry declares `Device: Desktop|Lenovo` and a `Required On Other Device` field, and
`mutual_command_log_monitor.py` / `sync_onenote_command_log.ps1` / `append_onenote_command_log.ps1` read and
write it so the two machines can detect and apply each other's rule/script changes (`device-monitor`,
`mutual-check`). Device identity is auto-detected from `COMPUTERNAME` (or `AGENT_DEVICE_LABEL` override) in
`detect_device_label()` (duplicated in `agent_onenote_logger.py` and `mutual_command_log_monitor.py`).

### Journal pipeline (`summarize_note5.py`, ~2600 lines, the largest/central script)

Reads free-form notes from `rawtext/` (or a specific file), and either summarizes them via the OpenAI API or
(with `--local-summary`) extracts structure without any external call. Supports `--watch`/`--poll` for continuous
processing, `--delta-only` for incremental notes, and a small in-note command language ("研究OS v1.0"): `@todo`,
`@python` (sandboxed to `.py` files inside the workspace), `@ask`, `@命令`/`＠命令`, executed via
`--execute-commands` / `--commands-only`. Output journals go to `日誌/` and always end with a `## 生データ` raw-text
section (generation timestamp + rawtext mtime) so verification/debugging can diff against the source. TODOs
extracted with `@todo` can sync to Google Tasks (`--sync-google-todos`, gated by `.env`'s `GOOGLE_TODO_SYNC`) or,
in local/manual mode, get written to a "Google Tasks手動投入用" copy-paste TodoList instead of being sent anywhere.

### Other scripts worth knowing about

- `agent_file_logger.py` — shared low-level Markdown/JSON log writer used by the OneNote logger and conversation
  log agent (`AGENT_LOG_ROOT` = `agent_workspace/`).
- `conversation_log_agent.py` — appends structured conversation records (Markdown + JSONL) under
  `agent_workspace/会話ログAgent/`; called automatically by the orchestrator after every routed request.
- `migration_check.py` — environment/dependency/credentials/OneNote-COM health check (`Check` dataclass, OK/WARN/
  FAIL/INFO), writes reports under `agent_workspace/司令塔Agent/summaries/`.
- `mutual_command_log_monitor.py` — Desktop/Lenovo diffing logic driving `device-monitor`/`mutual-check`.
- `指示作業Onenote分類.py` / `指示作業分類.py` — OneNote page classification agents, keyed by
  `tools/onenote_classification_maps/*.json`.
- `paper_finding.py`, `daily_paper_search.py`, `paper_to_onenote.py` — Crossref + OpenAI-scored literature search,
  writing `papers/YYYY-MM-DD_paper_list.xlsx`.
- `sync_google_todo_queue.py`, `setup_google_tasks.py` — Google Tasks queue sync and OAuth setup
  (`日誌/google_todo_queue.jsonl` → `日誌/google_todo_synced.jsonl`).
- `tools/*.ps1` — OneNote-COM-based agents (`onenote_search_agent.ps1`, `paper_search_agent.ps1`,
  `organic_synthesis_agent.ps1`) called by the orchestrator's routed commands.

### Directory map (generated/state directories, not hand-edited)

- `rawtext/` — raw input notes consumed by the journal pipeline.
- `日誌/` — generated daily journal Markdown, TODO/command logs, Google Tasks queues.
- `agent_workspace/<Agent名>/` — per-agent logs, JSON records, and state (one subfolder per agent: 司令塔Agent,
  OneNote検索Agent, 会話ログAgent, 実験ノートAgent, 検証Agent, etc.).
- `tools/` — the PowerShell OneNote-COM agents plus `orchestrator_agent_log.jsonl` (full routing history) and
  OneNote classification maps.
- `.note_agent_state/`, `work/` — misc. per-run state/output for ad-hoc tests.
- `papers/`, `paper/`, `summarize_paper/` — literature search outputs and summarized-paper artifacts.

## Conventions / rules to preserve

- **Route user-facing work through `agent.ps1`/the orchestrator**, not the underlying script directly, so
  verification, the Agent council, and the cross-device command log all fire. If a direct script run is
  unavoidable, immediately follow with `.\agent.ps1 log-command -Device Desktop -Summary "..."`.
- **Command-log entries** (OneNote `OpenAI_Agent1` / `命令したLog`) must keep the `## Summary / Actions / Files /
  Verification / Required On Other Device / Next steps` structure, with `Device` set to `Desktop` or `Lenovo` and
  file paths given relative to `OpenAI-Agent` — this is what keeps the two machines' automation in sync.
- **DOI/source requirements**: `論文検索Agent` and `有機合成Agent` outputs are verified to require a DOI or an
  explicit "DOI未発見" with documented search scope — don't silently drop this when touching those code paths.
- **Never paste `.env`, the Google OAuth client secret, or token file contents into chat** — `credentials.json`,
  `token_google_tasks.json`, `client_secret*.json`, and `*oauth*.json` are gitignored for this reason.
- Most scripts add `.venv/Lib/site-packages` and/or the bundled Codex runtime's `site-packages` to `sys.path`
  manually near the top of the file (see `summarize_note5.py`'s `site_packages_compatible` check) — this is
  intentional because the checked-in `.venv` may have compiled extensions built for a different Python ABI than
  whatever interpreter actually runs the script.
