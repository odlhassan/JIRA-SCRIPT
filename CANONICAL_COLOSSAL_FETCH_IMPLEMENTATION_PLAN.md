# Canonical Colossal-Fetch Implementation Plan

## Summary

This project introduces a single canonical yearly Jira refresh that populates one shared SQLite dataset for all reports.

## Two-Phase Refresh Lifecycle (2026-08-28)

Colossal Refresh now separates data acquisition from precomputation. A global Fetch writes an immutable raw canonical snapshot; a global Compute rebuilds the shared derived tables, compatibility artifacts, and report output from one successful Fetch snapshot. Reports continue to use the previously promoted precomputed version until Compute succeeds.

- `POST /api/canonical-refresh` remains the one-click Fetch then Compute workflow.
- `POST /api/canonical-fetch` performs Fetch only; `POST /api/canonical-compute` computes the latest (or requested) successful Fetch.
- Smart Fetch uses a durable checkpoint with a ten-minute overlap. A Smart request due for reconciliation upgrades to a full inventory reconciliation after seven days.
- `canonical_fetch_runs` and `canonical_compute_runs` record independent lifecycle status. Existing `canonical_refresh_runs` and `last_success_run_id` remain the report compatibility contract.
- Employee Performance per-assignee refresh is deliberately scoped: it records `employee_performance_scoped_runs`, regenerates only Employee Performance output, and never promotes its partial snapshot as the global canonical/report version.
- Completed lifecycle history is pruned after 30 days without SQLite compaction; active/current and fallback snapshots are retained.

The refresh is initiated from the Nested Report UI and is scoped only to active projects from the Managed Projects module.

The end-state design is:

- one shared canonical fetch
- one canonical SQLite source of truth
- reports compute their own slicing/dicing from canonical data
- legacy XLSX outputs are only temporary compatibility bridges during migration

This plan is informed by:

- [EMPLOYEE_PERFORMANCE_REFRESH_AUDIT.md](E:/JIRA%20SCRIPT/EMPLOYEE_PERFORMANCE_REFRESH_AUDIT.md)

## Locked Rules

- Project scope comes only from Managed Projects.
- Only active managed projects are fetched.
- Employee data is only fetched insofar as it belongs to those managed projects.
- Refresh scope is year-based.
- Inclusion rules are:
  - start in year, or
  - due in year, or
  - updated in year, or
  - worklog in year
- Parent/child hierarchy completion is mandatory.
- Worklogs are stored only from the selected start month through December 31 of the selected year. January remains the backward-compatible default.
- Work proceeds phase-by-phase.
- After each completed phase, pause and request permission before starting the next phase.

## Phase 1 — Canonical Refresh Foundation

### Goal

Introduce the canonical refresh contract, schema, and lifecycle APIs without migrating report consumers yet.

### Scope

- Add canonical refresh subsystem in the server layer
- Add canonical schema tables
- Add run/state tracking
- Add canonical refresh APIs
- Use Managed Projects as the project-scope authority

### Canonical tables

- `canonical_refresh_runs`
- `canonical_refresh_state`
- `canonical_issues`
- `canonical_issue_links`
- `canonical_worklogs`
- `canonical_issue_scope_reasons`
- `canonical_sync_state`

### API surface

- `POST /api/canonical-refresh`
- `GET /api/canonical-refresh/<run_id>`
- `GET /api/canonical-refresh/current`
- `POST /api/canonical-refresh/cancel`

### Status

Implemented.

## Phase 2 — Colossal Fetch Engine

### Goal

Implement the actual yearly Jira ingestion into canonical storage.

### Scope

- Load active managed project keys at refresh start
- Resolve year window
- Discover candidate issues using four inclusion signals
- Expand hierarchy upward and downward
- Fetch full issue details
- Fetch worklogs for included issues
- Persist canonical issue/link/worklog/scope-reason rows
- Record run stats
- Support async progress and cancel

### Current implementation behavior

The canonical refresh now:

1. Loads active managed projects
2. Resolves the selected start month and year-end boundaries
3. Discovers issues by:
   - planned date in scope
   - updated in scope
   - worklog in scope
4. Expands parents and descendants needed for hierarchy completeness
5. Fetches detailed issue payloads
6. Fetches worklogs for included issues
7. Writes canonical rows into SQLite
8. Records scope reasons and run stats

### Status

Implemented.

## Phase 3 — Derived Shared Read Models

### Goal

Create reusable shared structures so reports do not each re-derive heavy logic independently.

### Scope

- Build derived/materialized tables or cached views for:
  - epic/story/subtask hierarchy view
  - actual date rollups from worklogs
  - assignee daily/weekly/monthly aggregates
  - planning completeness flags
  - employee-performance-oriented slices
  - project/assignee summary helpers used by multiple reports
- Ensure derivation only uses canonical rows for managed projects in the selected run
- Centralize shared logic for downstream reports

### Status

Implemented.

## Phase 4 — Compatibility Bridge from Canonical DB

### Goal

Keep existing reports working while migration is in progress by rebuilding legacy artifacts from canonical data instead of Jira.

### Scope

- Regenerate temporary compatibility artifacts from canonical DB:
  - `1_jira_work_items_export.xlsx`
  - `2_jira_subtask_worklogs.xlsx`
  - `3_jira_subtask_worklog_rollup.xlsx`
  - `nested view.xlsx`
  - any additional still-required legacy artifacts
- Ensure bridge artifacts include only managed-project data
- Ensure bridge generation performs no Jira calls

### Status

Implemented.

## Phase 5 — Report Migration Group 1

### Goal

Migrate the file-backed reports with the clearest canonical overlap first.

### Reports

- Nested View
- Dashboard
- Missed Entries

### Scope

- Replace file-based reads with canonical DB / derived-view readers
- Remove report-local Jira fetches
- Preserve existing filter semantics and output shape
- Keep report project choices aligned to Managed Projects

### Status

Implemented.

## Phase 6 — Report Migration Group 2

### Goal

Migrate the assignee and leave oriented reports.

### Reports

- Assignee Hours
- Employee Performance
- Leaves Planned Calendar
- RLT Leave Report if brought under the canonical boundary

### Scope

- Replace `epf_*` staging as primary fetched-source truth
- Rebuild these reports from canonical work items/worklogs plus existing settings/config tables
- Restrict employee-level calculations to managed-project data only

### Status

Implemented.

## Phase 7 — Report Migration Group 3

### Goal

Migrate the remaining hierarchy/API-driven reports.

### Reports

- Planned RMIs
- Gantt Chart
- Phase RMI Gantt
- Planned vs Dispensed
- Planned Actual Table View
- Original Estimates Hierarchy
- IPP-related consumers where canonical overlap is intended

### Scope

- Replace direct Jira hierarchy fetches with canonical hierarchy queries where appropriate
- Keep only explicit exceptions for truly specialized consumers
- Preserve Managed Projects as the sole project-scope source

### Status

Implemented, with `ipp_meeting_dashboard` still left as an explicit specialized exception.

## Phase 8 — Final Cutover and Cleanup

### Goal

Retire duplicate fetch systems and make canonical refresh the only shared fetched-data path.

### Scope

- Remove obsolete report-specific Jira refresh paths
- Remove temporary bridge artifacts once all consumers are migrated
- Keep export generation only where needed for user-facing export
- Add integrity tooling and audit support
- Update UI language to reflect canonical refresh semantics

### Status

Not started.

## Current Code Status

Implemented so far:

- canonical schema initialization
- canonical refresh lifecycle APIs
- managed-project-scoped async canonical refresh
- yearly issue discovery
- hierarchy expansion
- issue detail fetch
- worklog fetch
- canonical persistence
- canonical derived read-model rebuild
- issue actual-date rollups
- assignee day/week/month aggregates
- planning completeness flags
- hierarchy summary cache
- project/assignee summary helpers
- canonical-to-legacy compatibility workbook bridge
- canonical hierarchy-backed Planned vs Dispensed loader
- canonical-backed Planned Actual Table View refresh path
- canonical-backed Original Estimates refresh path
- canonical-backed report refresh chains for:
  - Planned RMIs
  - Gantt Chart
  - Phase RMI Gantt
  - Planned vs Dispensed
  - Planned Actual Table View
  - Original Estimates Hierarchy
- `1_jira_work_items_export.xlsx` regeneration from canonical DB
- `2_jira_subtask_worklogs.xlsx` regeneration from canonical DB
- `3_jira_subtask_worklog_rollup.xlsx` regeneration from canonical DB
- `nested view.xlsx` regeneration from canonical DB
- scope-reason tracking
- run stats recording
- Group 1 report refreshes moved off report-local Jira fetch
- Missed Entries snapshot/HTML rebuild from canonical DB
- Dashboard refresh reuses canonical-backed artifacts only
- Nested View actual-hours API prefers canonical DB
- Dashboard actual-hours aggregate API prefers canonical DB
- Group 2 report refreshes moved off Jira fetch and onto canonical-backed rebuilds
- `generate_rlt_leave_report.py` supports canonical DB source mode
- `generate_employee_performance_report.py` supports canonical DB source mode for work items/worklogs
- Assignee Hours refresh now rebuilds from canonical compatibility artifacts plus canonical leave report
- Employee Performance async refresh uses canonical rebuild flow instead of `epf_*` staging when canonical data exists
- Leaves Planned Calendar refresh now rebuilds from canonical leave report
- RLT Leave Report refresh now rebuilds directly from canonical DB
- **Incremental sync-cache is now warmed per-issue inside the worklog fetch loop** — each issue+worklog pair is written to `jira_sync_cache.db` immediately after being fetched, so a subsequent smart refresh after a failed/cancelled run will skip already-fetched issues automatically (resume behaviour)
- **Resume (Smart Refresh) banner** added to the Colossal Refresh settings UI — shown when the most recent run was `canceled` or `failed`; clicking it pre-fills the scope year from the last run and starts a smart refresh
- **Azure split-path output fix** — when `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH` points to persistent storage such as `/home/data/assignee_hours_capacity.db`, canonical data is stored in that DB but compatibility workbooks, generated report HTML, and `report_html` sync still run from the application root. This prevents successful Colossal Refresh runs from writing fresh artifacts under `/home/data` where the live site does not serve them.
- **Epics Planner sync from canonical** — new `syncing_epics_management` stage runs after `rebuilding_derived_data`. For each Epic-type issue in the current canonical run, any matching row in `epics_management` (matched by `epic_key`) has its `epic_name`, `start_date`, `due_date`, and `jira_url` updated with the latest Jira values. Only non-user-managed fields are touched; plan fields, delivery status, and plan JSON are never overwritten. This ensures `settings/product-releases` shows current epic names and dates after every Colossal Refresh.

Key implementation location:

- [report_server.py](E:/JIRA%20SCRIPT/report_server.py)

Current tests added:

- [tests/test_canonical_refresh_api.py](E:/JIRA%20SCRIPT/tests/test_canonical_refresh_api.py)
- [tests/test_group2_canonical_refresh.py](E:/JIRA%20SCRIPT/tests/test_group2_canonical_refresh.py)

## Business Logic

Colossal Refresh writes canonical Jira issue, hierarchy, worklog, and derived read-model tables to the configured capacity DB. Reports select the latest promoted successful run from `canonical_refresh_state.last_success_run_id`; a refresh is visible only after the run status becomes `success` and that state pointer is updated. On Azure, the DB path may be external (`/home/data`) while generated artifacts must remain under the app root so served static reports use the latest output.

## Business Cases

The module gives the EPR Tool a single Jira fetch for all managed projects, avoiding report-by-report Jira refresh drift. Business users use the output for dashboards, Monthly Epic Plan vs Actual, employee performance, RLT leave reporting, nested views, and planning reports.

## Examples

If Azure has `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db`, a 2026 Colossal Refresh stores `canonical_issues` and `canonical_worklogs` in `/home/data/assignee_hours_capacity.db`, then rebuilds `1_jira_work_items_export.xlsx`, `nested view.xlsx`, generated reports, and `report_html` from the app root so `https://epreporting.azurewebsites.net/...` serves the fresh artifacts.

## Explanations

The refresh loads active managed projects, discovers Jira issues from the selected start month through December 31 of the selected year, expands parent/child hierarchy, fetches details and in-scope worklogs, stores the result as a canonical run, rebuilds derived tables, regenerates compatibility artifacts for legacy consumers, and syncs generated HTML into the served report folder.

## Front-end UI Fields

The Colossal Refresh settings page exposes the selected year and start month, refresh mode (`full` or `smart`), start/resume/cancel controls, optional full-refresh database backup, active run progress, stage status, and the latest successful run summary. Report screens then read the latest promoted run without asking users to select a run id.

| Field | Type | Default | Applies To | Side Effect |
| --- | --- | --- | --- | --- |
| Scope Year | Number input | Current year | Smart and Full Refresh / Fetch Only | Sets the year whose December 31 date closes the canonical Jira window. Valid range: 2000-2100. |
| Start Month | Select (January-December) | January | Smart and Full Refresh / Fetch Only | Sets the first day included in issue discovery and worklog storage. The scope runs from day 1 of this month through December 31. Resume restores both year and month; Smart Refresh only reuses a successful baseline with the same year, month, and managed-project set. |
| Replace previous fetched Jira data and reclaim disk space | Checkbox | Checked | Complete Full Refresh only | After successful report generation and snapshot promotion, deletes all other canonical snapshots (including other date windows) and attempts SQLite VACUUM. Uncheck to retain history. Smart Refresh and standalone Fetch ignore this option. Settings and manual plans remain. |
| Full database backup | Not shown | Disabled | All refresh modes | Multi-gigabyte full-DB copies exhausted production disk, so the UI no longer presents this option. The backend accepts but ignores the legacy `create_db_backup` request flag. |

## Script Files

- `report_server.py` — Flask routes, canonical refresh lifecycle, artifact rebuild orchestration, and report HTML sync.
- `canonical_report_data.py` — shared canonical DB readers and latest-run resolution.
- `monthly_epic_plan_progress_service.py` — Monthly Epic Plan vs Actual payload builder from canonical tables.
- `generate_rlt_leave_report.py`, `generate_employee_performance_report.py`, `support_center_sync.py` — downstream generators invoked by the refresh.
- `tests/test_canonical_refresh_api.py` — API and Azure split-path regression coverage.
- `tests/test_group2_canonical_refresh.py` — downstream canonical refresh chain coverage.

## Dependent & Impacted Files

Reports and modules that consume the latest canonical run include dashboard/nested-view APIs in `report_server.py`, `monthly_epic_plan_progress_service.py`, generated report scripts, `report_html/*` served assets, and compatibility workbooks consumed by legacy report generators.

## Table Schema

Primary tables in `assignee_hours_capacity.db` include `canonical_refresh_runs` (run id, scope, status, timestamps, stats), `canonical_refresh_state` (active and last successful run pointers), `canonical_issues` (one row per Jira issue per run), `canonical_issue_links`, `canonical_worklogs`, `canonical_issue_scope_reasons`, derived `canonical_issue_actuals`, `canonical_assignee_period_hours`, `canonical_issue_planning_flags`, `canonical_hierarchy_summary`, and `canonical_project_assignee_summary`.

## Data Flow

1. User starts Colossal Refresh from `/settings/canonical-refresh`.
2. The UI sends `year` and `start_month` to `POST /api/canonical-refresh` or `POST /api/canonical-fetch`; the server rejects months outside 1-12 and defaults omitted months to January for existing clients.
3. If the user checked `Create DB backup before Full Refresh` and started a full refresh, `report_server.py` creates a timestamped backup under `backups/canonical_refresh`.
4. `report_server.py` writes run state and canonical rows into the configured capacity DB. The start month is retained in existing run metadata, so no SQLite schema migration is required.
5. The successful run updates `canonical_refresh_state.last_success_run_id`.
6. Derived tables and compatibility artifacts are rebuilt from that run.
7. Generated HTML is synced from the app root into `report_html`.
8. Live report routes and APIs read the latest successful canonical run and served artifacts.

### Worker Restart Recovery

Fetch completion is authoritative in `canonical_fetch_runs`. If the Flask worker ends after Fetch has committed its canonical issues and worklogs but before Compute finishes, `/api/canonical-refresh/current` presents the combined run as `fetch_ready` at the `fetch_done` step rather than discarding the snapshot. The page directs the user to **Finish Refresh — Build Reports**. `POST /api/canonical-compute` accepts the durable successful Fetch record even when the older combined `canonical_refresh_runs` row was marked `failed`. A successful retry promotes the legacy row back to `success` and updates the active report pointer. If a retrying Compute worker also ends, its stale running row is failed and cleared so Compute can be retried again from the same Fetch. If the worker ends before Fetch commits successfully, the Fetch and combined run are both failed and must be fetched again.

The current UI labels this recovery action **Finish Refresh — Build Reports**, explicitly says that Jira will not be fetched again, and shows later stages as **Waiting** rather than generic **Pending**.

Related regression coverage:

- [tests/test_missed_entries_refresh_api.py](E:/JIRA%20SCRIPT/tests/test_missed_entries_refresh_api.py)

## Verification Notes

The following commands have been used to verify current implemented phases:

```powershell
python -m pytest tests/test_canonical_refresh_api.py -q
python -m pytest tests/test_missed_entries_refresh_api.py -q
```

## Next Planned Phase

Next phase to start after approval:

- Phase 3 — Derived Shared Read Models

### Full Refresh snapshot replacement and disk reclamation

**Business logic and cases.** Administrators can replace a large historical Jira pull
with a smaller month/year scope without resetting the mixed settings/planning database.
The checkbox is checked on the settings page; API clients must explicitly send the
boolean `replace_previous_data: true` to `POST /api/canonical-refresh` with `mode: full`.
Omitting it preserves existing API behavior. Smart Refresh, including a smart run that
falls back to a full fetch, and standalone Fetch never request replacement.

**Example and explanation.** Select 2026 / October and leave the checkbox checked.
Fetch writes a new snapshot alongside the old one. Compute builds derived outputs,
synchronizes Epic metadata and generates reports. Only after success does it promote
the new snapshot, delete older canonical issue/link/worklog/reason/derived rows and
obsolete refresh/fetch/compute metadata, and run VACUUM under the existing refresh lock.
Old employee-scoped run metadata referencing discarded snapshots is removed too.
Failed or canceled builds keep the previously published snapshot. A manually retried
Compute retains the choice saved with its original Fetch. Cancellation is checked again
before promotion; once cleanup starts it is not cancelable from the page.

**Files, table schema and dependencies.** `canonical_storage.py` owns an explicit
allowlist of the nine canonical snapshot tables (all keyed by `run_id`) and cleanup
of lifecycle metadata described above. It never scans arbitrary tables for deletion.
`report_server.py` stores the choice and cleanup result in existing `stats_json` TEXT
columns; there are no table, column, index or constraint changes. No schema migration
is required. `tests/test_canonical_storage.py` covers cleanup, API scope, failure,
cancellation, low disk space, integrity and real file shrinkage. All canonical readers
in `canonical_report_data.py` and the report generators continue using the promoted
run pointer. Settings, capacity, teams, manual Epic plans, product releases, seating,
other report snapshot families, and the separate `jira_sync_cache.db` are preserved.
This action does not promise to clear every cache or make the DB equal the Jira pull size.

**Data flow and diagnostics.** The checkbox travels from the settings page to queued
run JSON, Fetch progress JSON, then successful Compute. Replacement skips the general
30-day history cleanup before promotion; replacement itself owns history removal.
Cleanup validates the published snapshot and other active jobs before deleting in
batches of at most 5,000 rows per transaction. A brief pause between full batches
allows status/report readers to acquire a lock on Azure's shared drive. Successful batches remain committed if
a later batch fails, and a later replacement can resume removing old data. Cleanup
repairs successful Fetch and Compute pointers after deletion. It refuses an unpublished
snapshot or another running job identified by its lifecycle's active pointer. Historical
rows still labeled `running` without an active pointer are abandoned history, and do
not block replacement under the refresh lock. The October 26 live retry exposed an
October 20 abandoned Fetch row with no active pointer. Compaction uses SQLite VACUUM in place, never
unlinks a live DB or journal. The result exposes `storage_cleanup.status`, per-table
`deleted_rows`, `bytes_before`, `bytes_after`, and `bytes_reclaimed` in legacy run and
Compute stats; the page shows reclaimed MiB or a visible maintenance warning.

**Operational limits.** Replacement temporarily stores old and new snapshots together.
Deletion also needs journal space; bounded transactions allow freed pages to be reused
without accumulating one journal for all history. Before VACUUM, the server estimates
temporary space as twice the live page count (`page_count - freelist_count`) times
`page_size`, reported as `compaction_estimated_bytes`. The still-bloated file size is
not the estimate: for example, 9 GiB with 100 MiB of live pages estimates 200 MiB.
This is an estimate, and SQLite can still reject compaction. With insufficient space
or a SQLite lock/error, valid reports stay successful and the page states that cleanup
or compaction is pending. Freed pages remain reusable. A later checked Full Refresh
retries cleanup and compaction; this is not a background automatic retry. If the worker
is interrupted, normal Compute recovery can be used to finish again. Backups in other
folders and unrelated cache databases are not deleted.

For refreshes retaining history, a SQLite error in general 30-day retention is a
maintenance warning in Compute stats and the page's completion detail; it no longer
discards completed reports. This does not repair an already corrupt database. The
October 20 production logs showed malformed-image and disk-I/O errors in this cleanup
path, while the October 26 check found a 9.1 GiB database on a 97%-full persistent share.
Any database recovery must preserve settings/manual plans and independently validate
integrity; never delete live WAL/SHM files or substitute a local development database.
Application startup applies the journal policy before all settings initializers and
explicitly closes that connection, allowing an inherited WAL database to switch to
DELETE before other startup connections can lock the pragma. This changes no schema.
Canonical progress, promotion and terminal Compute status writes allow up to 60 seconds
for a read lock and explicitly close their connections. A brief report reader must not
turn a completed build into an orphaned worker merely by exceeding SQLite's default
five-second write timeout. A real six-second shared-reader regression test covers this.
The cleanup connection requests a 64 MiB page cache and disables dirty-page spilling
for its bounded transactions, avoiding repeated network reads and early exclusive
locks; other connections retain their normal settings. The cache size is a target,
and one batch's dirty pages may exceed it until commit. This follows
[SQLite's cache-spill guidance](https://www.sqlite.org/pragma.html#pragma_cache_spill).
Azure's SQLite build was also observed with `secure_delete=ON`. Cleanup uses `FAST`
on its own connection to avoid zeroing discarded overflow pages before VACUUM;
normal connections keep their original policy. Until compaction succeeds, freed
pages can still contain old bytes, as with other SQLite logical deletions. See
[SQLite secure-delete modes](https://www.sqlite.org/pragma.html#pragma_secure_delete).
If pruning raises an operational lock or disk-I/O error, Compute retries cleanup up
to three times with a fresh connection and short backoff, without fetching Jira or
rebuilding reports again. Earlier committed batches are retained. The result records
`storage_cleanup.attempts`; persistent failures remain visible warnings. Corruption,
full-disk errors, and active-job validation failures are not automatically retried.

**Local verification.** Restart with `.\.venv\Scripts\python.exe run_server.py --no-sync --keep-cache`
and open `http://127.0.0.1:3000/settings/canonical-refresh`. This settings page is generated
by the server, so no HTML rebuild is needed. Run focused tests with
`.\.venv\Scripts\python.exe -m unittest tests.test_canonical_storage tests.test_canonical_two_phase tests.test_canonical_refresh_api tests.test_report_ui_smoke.ReportUiSmokeTests.test_canonical_refresh_settings_formats_timestamps_for_display`.

**Production verification — October 26, 2026.** Full Refresh was started with
`year=2026`, `start_month=10`, `mode=full`, and `replace_previous_data=true`.
The initial report build failed with `database or disk is full`; its successful
Fetch was reused after deployment rather than acquired again. The completed
Compute reported `success`, `done`, and `storage_cleanup.status=compacted`:

- Fetch: `canonical-1792992881-01eb0fbb`; 1,514 issues and 235 worklogs retained.
- Compaction: 9,708,490,752 bytes before, 61,825,024 after, 9,646,665,728 logical bytes removed (physical SMB allocation verified separately below).
- Exactly one canonical issue snapshot remains, using DELETE journal mode.
- Full SQLite `integrity_check` returned `ok`; `foreign_key_check` found zero violations.
- Before/after counts matched in all ten sampled settings/planning tables, including
  3,211 Epic plan values, 20 approved-date rows, nine releases, and 24 booking headers.
- Executed focused checks: 53 tests passed with
  `.\.venv\Scripts\python.exe -m unittest tests.test_canonical_storage tests.test_canonical_two_phase tests.test_canonical_refresh_api tests.test_db_journal_mode tests.test_azure_deploy_contract -q`.
- The latest local server was started with
  `.\.venv\Scripts\python.exe run_server.py --port 3118 --no-sync --keep-cache`;
  inspect `http://127.0.0.1:3118/settings/canonical-refresh`.

### Azure physical allocation verification

VACUUM alone left the original Azure SMB allocation attached to the compact file.
`canonical_storage.py` now uses an isolated, exclusively locked maintenance process
to punch the unused tail inside a temporary EOF extension, restoring the original
EOF and verifying SHA-256 of every original byte. Authoritative CIFS server
allocation is recorded as `allocated_bytes_before/after`; `bytes_reclaimed` uses
those values when available. Unsupported filesystems use ordinary VACUUM. Failure
leaves an explicit maintenance warning, rather than falsely claiming compaction.
Focused tests cover the helper boundary, physical-byte accounting, and warning path.
No schema changes or database inode replacement are involved.

Production verification on 2026-10-26: authoritative server allocation fell from
9,708,503,040 to 61,931,520 bytes. `/home` usage fell to 7%, with 9.4 GiB free.
Every original database byte matched its SHA-256 after trimming; full integrity
check returned `ok` and foreign-key check returned zero violations. A verified
compact backup was retained before maintenance. Both Linux helper safety tests
passed on Azure (including failed hole-punch restoring EOF and releasing locks),
and a throwaway SQLite DB exercised actual Azure physical allocation release.
The focused application suite passed 56 tests locally.

When report HTML promotion cannot write to Azure's read-only package, GET report
routes now serve the selected rebuilt source directly, even if an older packaged
`report_html` copy exists. Tests in `tests/test_report_html_promotion.py` cover both
read-only temp-file creation and failed atomic promotion. This prevents a successful
Compute from leaving the browser on stale packaged report content.
