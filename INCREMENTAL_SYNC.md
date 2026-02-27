# Incremental Jira Sync

This project supports SQLite-backed incremental fetch for:

- `export_jira_subtask_worklogs.py`
- `export_jira_work_items.py`
- `export_jira_nested_view.py`
- `generate_rlt_leave_report.py`

The scripts fetch only Jira issues changed since the last checkpoint, but still rebuild full Excel outputs from the local cache for deterministic downstream behavior.

## Environment variables

- `JIRA_SYNC_DB_PATH`  
  SQLite path. Default: `jira_sync_cache.db`

- `JIRA_INCREMENTAL_OVERLAP_MINUTES`  
  Overlap window to avoid misses. Default: `5`

- `JIRA_FORCE_FULL_SYNC_DAYS`  
  Force a full reconciliation every N days. Default: `7`

- `JIRA_INCREMENTAL_BOOTSTRAP_DAYS`  
  Initial fallback checkpoint age if no state exists. Default: `365`

- `JIRA_INCREMENTAL_DISABLE`  
  Set `1` to disable incremental mode (full sync behavior). Default: `1`

## CLI toggle

Incremental fetch is **opt-in** per run. Use the `--incremental` flag when running:

- `export_jira_subtask_worklogs.py --incremental`
- `export_jira_work_items.py --incremental`
- `export_jira_nested_view.py --incremental`
- `generate_rlt_leave_report.py --incremental`
- `run_all_exports.py --incremental`

## How it works

1. Read last checkpoint from `sync_state`.
2. Run discovery query on Jira:
   - `updated >= checkpoint - overlap` (or full query on forced full sync).
3. Detect new/changed issues using cached `issue_index`.
4. Fetch full details only for new/changed issues.
5. Update cache tables:
   - `issue_index`
   - `issue_payloads`
   - `worklog_payloads` (worklogs exporter only)
6. Rebuild complete Excel file from cache snapshot.
7. Advance checkpoint to max `updated` seen.

## Cache reset and recovery

To force a clean bootstrap:

1. Delete the SQLite file:
   - `jira_sync_cache.db` (or your custom `JIRA_SYNC_DB_PATH` path)
2. Run exports again:
   - first run will execute full sync and repopulate cache.

## Troubleshooting

- No rows but Jira has data:
  - Verify project keys in `JIRA_PROJECT_KEYS`.
  - Run without `--incremental` to force a full fetch.

- Suspected stale/missing issue:
  - Run without `--incremental` to force full fetch, or delete DB and rerun.

- API rate-limit spikes:
  - Use `--incremental` for smart fetch runs.
  - Increase `JIRA_WORKLOG_DELAY_SECONDS` for worklog-heavy runs.
