# Employee Performance Refresh Audit

## Question

When the Employee Performance dashboard header `Refresh` runs, is the fetched data sufficient to satisfy the other reports in this repo?

## Short answer

No.

The Employee Performance refresh collects only a subset of the shared Jira data pipeline and stores it in an Employee Performance specific snapshot (`epf_*` tables in `assignee_hours_capacity.db`). That is sufficient for rebuilding `employee_performance_report.html`, but it is not sufficient to refresh most other reports as they are currently designed.

There are two reasons:

1. The refresh writes data into Employee Performance specific storage, not into the shared artifacts other reports consume.
2. Several other reports require additional artifacts that the Employee Performance refresh never creates, such as:
   - `3_jira_subtask_worklog_rollup.xlsx`
   - `nested view.xlsx`
   - `assignee_hours_report.xlsx`
   - `team_rmi_gantt_items` sync outputs
   - direct API-driven Jira hierarchy fetches used by Planned vs Dispensed / Planned Actual / Original Estimates

## What the Employee Performance refresh actually does

Implementation: `report_server.py`, `_run_employee_performance_isolated_refresh(...)`

For refresh runs, the flow is:

1. Run `export_jira_subtask_worklogs.py`
2. Run `export_jira_work_items.py`
3. Run `generate_rlt_leave_report.py`
4. Load those temporary outputs
5. Insert into Employee Performance snapshot tables:
   - `epf_work_items`
   - `epf_worklogs`
   - `epf_leave_rows`
   - `epf_leave_issue_keys`
6. Run `generate_employee_performance_report.py` in DB mode using that snapshot

Important detail:

- The refresh uses temporary XLSX files inside a temp directory.
- It does not publish the normal shared outputs like `1_jira_work_items_export.xlsx`, `2_jira_subtask_worklogs.xlsx`, `rlt_leave_report.xlsx`, `assignee_hours_report.xlsx`, `nested view.xlsx`, or `3_jira_subtask_worklog_rollup.xlsx`.

So even where the raw Jira content overlaps, the refresh does not update the artifact locations that other reports expect.

## Shared refresh baseline in the repo

The broader shared export pipeline is `run_all_exports.py`, which normally produces:

1. `export_jira_subtask_worklogs.py`
2. `export_jira_work_items.py`
3. `export_jira_subtask_worklog_rollup.py`
4. `export_jira_nested_view.py`

That broader pipeline is the base for many reports.

The Employee Performance refresh only covers the first two exports plus leave generation, and even those are staged privately into `epf_*` snapshot tables instead of shared output files.

## Audit by report

### 1. Employee Performance

Status: Yes

Why:

- This refresh path is explicitly built for `generate_employee_performance_report.py`.
- That report can read from `epf_*` DB snapshots using:
  - `JIRA_EMP_PERF_INPUT_SOURCE=db`
  - `JIRA_EMP_PERF_RUN_ID=<run_id>`

Conclusion:

- Fully supported.

### 2. Assignee Hours

Status: No, not as currently implemented

Why:

- `generate_assignee_hours_report.py` expects:
  - `2_jira_subtask_worklogs.xlsx`
  - `1_jira_work_items_export.xlsx`
  - `rlt_leave_report.xlsx`
  - capacity settings DB
- Employee refresh does not publish those shared files.
- It also does not run `generate_assignee_hours_report.py`, so `assignee_hours_report.xlsx` is not refreshed.

Raw overlap:

- Partial. Employee refresh does fetch worklogs, work items, and leave data.

Operational outcome:

- Not sufficient for the Assignee Hours report without refactoring that report to read from `epf_*` tables or a shared normalized store.

### 3. RLT Leave Report

Status: No, not as currently implemented

Why:

- Employee refresh runs `generate_rlt_leave_report.py`, but only to create a temporary file for Employee Performance ingestion.
- It does not update the normal shared `rlt_leave_report.xlsx` artifact.
- Employee snapshot tables only keep leave rows and leave issue keys needed by Employee Performance, not the full workbook contract other leave-based reports assume.

Raw overlap:

- Partial to strong, but not published in the expected place/format.

### 4. Leaves Planned Calendar

Status: No

Why:

- `generate_leaves_planned_calendar_html.py` reads `rlt_leave_report.xlsx`.
- Employee refresh does not refresh that shared workbook.

### 5. Missed Entries

Status: No, but close in terms of raw data

Why:

- `generate_missed_entries_html.py` reads `1_jira_work_items_export.xlsx`.
- Employee refresh does load equivalent work-item data into `epf_work_items`.
- However, the report is not implemented to read `epf_work_items`; it expects the standard workbook.

Raw overlap:

- Strong.

Operational outcome:

- Not sufficient without adapting Missed Entries to read from the Employee snapshot or a shared normalized store.

### 6. Dashboard

Status: No

Why:

- `fetch_jira_dashboard.py` reads:
  - `1_jira_work_items_export.xlsx`
  - `2_jira_subtask_worklogs.xlsx`
  - `3_jira_subtask_worklog_rollup.xlsx`
  - planner DB data
- Employee refresh never creates the rollup file.
- It also does not publish refreshed shared work-item/worklog artifacts.

Raw overlap:

- Partial.

Missing critical dependency:

- `3_jira_subtask_worklog_rollup.xlsx`

### 7. Nested View

Status: No

Why:

- `generate_nested_view_html.py` depends on `nested view.xlsx`, work items, and leave-related inputs.
- Employee refresh never runs `export_jira_nested_view.py`.
- Therefore `nested view.xlsx` is not refreshed.

### 8. Planned RMIs

Status: No

Why:

- `generate_planned_rmis_html.py` depends on:
  - `nested view.xlsx`
  - `1_jira_work_items_export.xlsx`
- Employee refresh does not create `nested view.xlsx`.

### 9. Gantt Chart

Status: No

Why:

- `generate_gantt_chart_html.py` depends on:
  - `nested view.xlsx`
  - `1_jira_work_items_export.xlsx`
- Employee refresh does not generate `nested view.xlsx`.

### 10. Phase RMI Gantt

Status: No

Why:

- `generate_phase_rmi_gantt_html.py` depends on synced team RMI gantt data in `assignee_hours_capacity.db`.
- That data is built via `sync_team_rmi_gantt_sqlite.py`, which consumes `1_jira_work_items_export.xlsx`.
- Employee refresh does not run that sync step.

### 11. IPP Meeting Dashboard

Status: No

Why:

- It depends on `export_ipp_phase_breakdown.py` inputs and its own transformed workbook flow.
- Employee refresh does not touch that pipeline.

### 12. Planned vs Dispensed

Status: No

Why:

- This report is API-driven.
- In `report_server.py`, `_load_planned_vs_dispensed_hierarchy(...)` fetches Jira hierarchy directly from Jira using JQL and project/date filters.
- It does not use the Employee Performance snapshot tables.

Raw overlap:

- Some overlapping issue fields exist, but the report is architected around direct API-backed hierarchy fetches, not Employee Performance staged data.

### 13. Planned Actual Table View

Status: No

Why:

- This report is API-driven and DB-backed through its own service/storage path.
- It is not wired to consume `epf_*` snapshot tables.

### 14. Original Estimates Hierarchy

Status: No

Why:

- This report is also API-driven and DB-backed through its own refresh path.
- It does not reuse Employee Performance snapshot storage.

## Summary table

| Report | Sufficient from Employee Performance refresh? | Main reason |
|---|---|---|
| Employee Performance | Yes | Native consumer of `epf_*` snapshot tables |
| Assignee Hours | No | Expects shared XLSX artifacts and its own summary generation |
| RLT Leave Report | No | Temp leave output is not published as shared artifact |
| Leaves Planned Calendar | No | Expects shared `rlt_leave_report.xlsx` |
| Missed Entries | No | Raw overlap exists, but report expects `1_jira_work_items_export.xlsx` |
| Dashboard | No | Needs rollup plus shared artifacts |
| Nested View | No | Needs `nested view.xlsx` |
| Planned RMIs | No | Needs `nested view.xlsx` |
| Gantt Chart | No | Needs `nested view.xlsx` |
| Phase RMI Gantt | No | Needs `sync_team_rmi_gantt_sqlite.py` outputs |
| IPP Meeting Dashboard | No | Separate export pipeline |
| Planned vs Dispensed | No | Direct Jira API hierarchy fetch path |
| Planned Actual Table View | No | Separate API/DB-backed refresh path |
| Original Estimates Hierarchy | No | Separate API/DB-backed refresh path |

## Final conclusion

The Employee Performance refresh is not a general-purpose shared data refresh.

It is a report-specific refresh that:

- fetches overlapping raw Jira data
- reshapes it into Employee Performance specific snapshot tables
- rebuilds only the Employee Performance report

If the goal is to make one refresh power multiple reports, the current Employee Performance refresh is the wrong system boundary.

## Recommended design direction

If you want one refresh to serve multiple reports, the better design is:

1. Build one shared normalized refresh layer
   - shared DB tables for work items, worklogs, leave data, rollups, nested hierarchy
2. Make all reports read from that shared store
   - instead of each report expecting its own XLSX contract
3. Treat report-specific snapshots as optional caches
   - not as the source of truth
4. Keep API-driven reports aligned on the same normalized cache when possible
   - or clearly isolate them if they must stay direct-to-Jira

In short:

- Employee Performance refresh data is sufficient for Employee Performance.
- It is not sufficient for all other reports in the current architecture.
