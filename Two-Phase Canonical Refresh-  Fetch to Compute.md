# Two-Phase Canonical Refresh: Fetch to Compute

## Summary

Split global Colossal Refresh into independent Fetch and Compute phases. Reports continue reading the current shared precomputed database version. The default remains one click: Fetch, then auto-Compute, with advanced phase-specific actions.

Employee Performance has one exception: its per-assignee refresh remains a separate scoped Fetch to Employee Performance-only Compute pipeline and must never promote partial data as the global canonical/report version.

## Core Changes

- Add lifecycle/state records for:
  - Global Fetch runs: mode, scope year, managed-project fingerprint, checkpoint, baseline, reconciliation state, status, and audit statistics.
  - Global Compute runs: source Fetch run, output version, status, report-generation results, timestamps, and errors.
  - Employee Performance scoped refresh runs: assignee, scope, source state, Fetch status, Employee Performance Compute status, and output version.
- Extend canonical state with independent global Fetch and Compute pointers plus `last_full_reconciliation_at_utc`.
- Keep raw canonical issue/link/worklog data keyed by global Fetch run; keep global derived/precomputed rows keyed by global Compute run.
- Preserve Employee Performance's dedicated scoped snapshot/output so an assignee refresh updates only that report.

## Fetch and Compute Behavior

- Global Full Fetch:
  - Fetch complete managed-project/year scope, expand hierarchy, fetch in-scope worklogs, validate, and persist a complete raw snapshot.
- Global Smart Fetch:
  - Use a stored successful checkpoint minus a 10-minute overlap.
  - Fetch changed/new Jira entities, replace each affected issue's complete in-scope worklog set, apply scope rules, expand affected hierarchy, and merge with the prior complete Fetch snapshot.
  - Force Full Fetch when project scope or year changes, no valid baseline exists, or reconciliation is due.
- Due-based reconciliation:
  - On the first Fetch started seven or more days after the prior successful reconciliation, upgrade Smart to Full Reconciliation.
  - Compare full Jira inventory/worklogs against the previous snapshot to remove deleted or missed records.
- Global Compute:
  - Accept a successful Fetch run and never call Jira.
  - Rebuild all global derived tables, compatibility artifacts, Epics Planner Jira-owned fields, generated reports, and served output.
  - Promote the global precomputed version only after complete validation and successful output synchronization.
  - A failed Compute leaves reports on the prior precomputed version.
  - The generators Compute runs (`generate_rlt_leave_report.py`, `generate_employee_performance_report.py`,
    `support_center_sync.py`) resolve their output directory through
    `report_output_paths.resolve_output_base()`, which mirrors
    `report_server._canonical_bridge_artifact_base_dir()`. Locally that is the repo root; on Azure,
    where `WEBSITE_RUN_FROM_PACKAGE` mounts `/home/site/wwwroot` read-only, it is the writable
    artifact directory the compatibility bridge uses (`$HOME/data/canonical_artifacts`, or
    `JIRA_CANONICAL_ARTIFACT_DIR`). Without this, Compute failed at `generate_rlt_leave_report.py`
    with `OSError: [Errno 30] Read-only file system`. Note that generators anchor paths to their own
    source directory, so setting the subprocess cwd alone does not fix this.
  - Promotion into `report_html` at the end of Compute is best-effort. A read-only publish target
    is recorded as `reports.sync_report_html_error` in the run stats and does not fail a Compute
    whose derived data and canonical pointers were already rebuilt.
- Employee Performance:
  - Global Compute regenerates Employee Performance from the global Fetch snapshot.
  - Per-assignee refresh performs an assignee-scoped Fetch followed by Employee Performance-only Compute.
  - It updates Employee Performance's scoped precomputed report data only; it never changes global Fetch/Compute pointers or any other report's data.

## APIs and UI

- Add:
  - `POST /api/canonical-fetch`
  - `POST /api/canonical-compute`
  - Fetch/Compute status endpoints with independent progress and cancellation.
- Preserve `POST /api/canonical-refresh` as backward-compatible orchestration: global Fetch followed automatically by global Compute.
- Keep `/api/employee-performance/assignee-refresh`; update its response/status to identify its scoped Fetch and Employee Performance-only Compute phases.
- Update Colossal Refresh settings with:
  - primary Smart Refresh and Full Refresh buttons;
  - advanced Fetch-only, Compute-latest, and recompute actions;
  - latest global Fetch, current global Compute/report version, reconciliation due status, and pending-compute visibility.
- Update Employee Performance UI only to show the scoped two-phase status; retain its existing assignee refresh workflow.

## Retention, Migration, and Tests

- Retain completed global Fetch, global Compute, and Employee Performance scoped runs for 30 days.
- Never prune active/current runs, their referenced source snapshots, or the latest successful fallback version.
- Update the SQLite schema changelog, schema snapshot, migration tool, migration documentation, and tests. Production migration remains pending until a production DB file is provided.
- Verify:
  - Smart delta, overlap, hierarchy movement, worklog add/edit/delete, scope changes, and reconciliation deletion cleanup.
  - Global Compute uses only a specified Fetch snapshot and preserves the prior report version on failure.
  - Per-assignee Employee Performance refresh updates that report while global dashboard, planning, nested, and other reports remain unchanged.
  - Existing report outputs continue to match between summary and drill-down views.

## Assumptions

- Global reports consume the current shared precomputed database. Employee Performance additionally supports a scoped assignee-refresh workflow, which is an exception that performs Employee Performance-only Fetch and Compute without changing the shared global canonical snapshot.
- The selected year and active managed projects remain the authoritative global scope.
- Automatic Fetch to Compute is the default; phase-specific actions are advanced controls.
- Smart Fetch uses a 10-minute overlap window.
- Weekly reconciliation is due-based: it runs on the first Fetch started at least seven days after the last successful reconciliation, not as an unreliable in-process timer.
- Retention is 30 days for completed Fetch and Compute history.

## Colossal Refresh UI behavior

The page groups related controls by purpose. **Refresh scope** contains the year and start month shared by all refresh actions. **Complete refresh** contains **Smart Refresh (recommended)** and **Full Refresh**; both fetch Jira, rebuild derived data, generate dependent reports, and publish served HTML. **Advanced workflow: run Fetch and report build separately** nests the phase-specific actions as two child groups: **1. Fetch Jira data only** and **2. Build and publish reports**. Operational buttons are grouped under **Run controls**, while offline preparation is separate under **Export tools**.

The stage display is split into two named groups:

1. **Fetch Jira data** — scope discovery, hierarchy expansion, issue/worklog retrieval, and durable canonical persistence.
2. **Build and publish reports** — derived data, Epics Planner synchronization, compatibility artifacts, dependent reports, and served HTML.

When Fetch is durable but Compute has not succeeded, the second group uses **Waiting** instead of the ambiguous **Pending** label. The page shows **Action required**, keeps overall progress at 85%, explains that currently published reports still use the prior successful version, and offers **Finish Refresh — Build Reports**. This action reuses the saved Fetch and does not call Jira again.

If a Compute retry failed, `GET /api/canonical-refresh/current` returns the newest Compute attempt for the latest Fetch, including its error. This lets the page explain the failure and offer a retry. Older one-phase databases whose successful run id is stored in both Fetch and Compute state pointers are treated as complete and are not incorrectly shown as requiring Compute.

After any successful Compute—automatic or manually started—the combined `canonical_refresh_runs` row is finalized as `success`, `done`, and `100%`. Its stage payload is also merged with the Compute results so API clients, persisted state, and the page all report the same completed lifecycle.

The full-database backup checkbox is no longer shown. Production full-DB copies were disabled because multi-gigabyte copies exhausted disk space, and the backend continues to ignore the backward-compatible `create_db_backup` request flag.
