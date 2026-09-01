# RLT Leave Report

Report ID: `rlt_leave_report`

INFO_IDS: `rlt.total_taken`, `rlt.total_planned_leaves`, `rlt.future_planned`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Total Taken | Leave already consumed in selected range. | `Planned Taken + Unplanned Taken` | planned taken hours, unplanned taken hours | Days are derived from configured daily-hours logic. | Nested leave-adjusted capacity and RnD capacity baseline. |
| Total Leaves Planned | Total planned leave load in selected range (taken + not-yet-taken). | `Planned Taken + Future Planned` | planned taken hours, planned not-yet-taken hours | Uses the same filtered range as RLT scorecards. | Nested total capacity adjusted and capacity gap. |
| Future Planned | Planned leave not yet consumed. | `Sum(planned estimates not yet taken)` | planned not-yet-taken hours | Missing required leave metadata is tracked in No Entry. | Nested adjusted capacity and assignee leave totals. |

## Drawer Notes

- Drawer describes leave categories and how they are reused in capacity-linked reports.

## Verification Signals

- The report uses `canonical_issues` and `canonical_worklogs` from `assignee_hours_capacity.db` by default, scoped to the last successful canonical refresh run and the RLT project. Direct Jira loading is a legacy, explicit `--source jira` option.
- The Verification Signals table and Excel sheet include both `start_date` and `due_date` alongside the created timestamp and derived verification reference date, so reviewers can compare the source leave window with the late-creation signal.

## Script Files

- `generate_rlt_leave_report.py` — generator. Writes `rlt_leave_report.xlsx`, `rlt_leave_report.html`, and `RLT_LEAVE_REPORT.md` relative to the process working directory. The workbook is written atomically (temp file, then `os.replace`); if the target directory rejects temp-file creation the workbook is staged in the system temp directory instead, and an unperformable final replace warns rather than raising.
- `canonical_report_data.py` — `build_rlt_leave_snapshot()` reuses the generator's aggregation helpers for API-driven reads.
- `report_server.py` — invokes the generator during Colossal Refresh Compute and promotes `rlt_leave_report.html` into `report_html`.
- `tests/test_rlt_leave_report.py`, `tests/test_readonly_root_report_generation.py` — coverage.

## Dependent & Impacted Files

- `report_server.py` — `REPORT_REFRESH_CHAINS` puts `generate_rlt_leave_report.py` ahead of `nested_view`, `assignee_hours`, `rnd_data_story`, `leaves_planned_calendar`, and `employee_performance`, so a failure here blocks those rebuilds. Both Colossal Refresh Compute paths (`_run_canonical_phase1_refresh`, `_run_canonical_compute`) run it as a subprocess.
- `_resolve_script_cwd()` in `report_server.py` decides where this generator's outputs land. On a writable app root it is the repo root; on Azure's read-only `WEBSITE_RUN_FROM_PACKAGE` mount it is `$HOME/data/canonical_artifacts` (or `JIRA_CANONICAL_ARTIFACT_DIR`). See `AZURE_APP_SERVICE.md`.
- `generate_employee_performance_report.py`, `generate_assignee_hours_report.py`, `generate_nested_view_html.py`, `generate_leaves_planned_calendar_html.py`, `generate_rnd_data_story.py` — downstream consumers of the leave outputs.

## Table Schema

Reads (from `assignee_hours_capacity.db`, scoped to `canonical_refresh_state.last_success_run_id` unless `JIRA_CANONICAL_RUN_ID` overrides it):

| Table | Columns used | Meaning |
| --- | --- | --- |
| `canonical_issues` | `run_id`, `issue_id`, `issue_key`, `project_key`, `issue_type`, `summary`, `status`, `assignee`, `start_date`, `due_date`, `created_utc`, `updated_utc`, `original_estimate_hours`, `total_hours_logged`, `parent_issue_key` | RLT leave subtasks and their parents for the snapshot. |
| `canonical_worklogs` | `run_id`, `worklog_id`, `issue_key`, `project_key`, `worklog_author`, `issue_assignee`, `started_date`, `hours_logged` | Actual leave hours consumed, used for taken vs planned classification. |
| `canonical_refresh_state` | `last_success_run_id` | Selects the active snapshot when no run id is supplied. |

The generator does not write to SQLite.
