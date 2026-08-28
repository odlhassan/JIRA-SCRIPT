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
