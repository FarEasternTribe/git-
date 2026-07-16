# Agent operation policy

## Automatically allowed

- Read-only inspection, search, status display, validation, and dry-runs.
- Local generated files under `agent_workspace`, `日誌`, `papers`, and runtime state folders.
- Idempotent append/sync operations that have a deduplication key and verification step.
- API summarization only when explicitly requested with `--api-summary` (journal actions are local-only by default).

## Explicit confirmation required

- Restoring or overwriting existing source/configuration files.
- Deleting files, clearing state, bulk rebuilding OneNote, or using a force option that bypasses deduplication.
- Moving or replacing credentials, starting a new OAuth authorization, or changing an external destination.
- Bulk external publication, email/message sending, purchases, or operations affecting third parties.
- Executing arbitrary code outside this workspace or expanding filesystem/network permissions.

An explicitly named launcher action counts as confirmation only for that action. Natural-language requests routed through `task` do not authorize destructive operations. Restore is dry-run by default and requires both `--apply` and `--confirm-restore`.

## Required controls

1. Use `agent.ps1` as the user-facing entry point.
2. Record each run in `.agent_runtime/state.json` with running/succeeded/failed status.
3. Deduplicate repeat mutating actions; bypass only when the user explicitly requests force.
4. Keep API keys, OAuth client files, and tokens outside the workspace and Git.
5. Verify important outputs before reporting success.
6. Create a secret-free configuration backup before a risky restore.

