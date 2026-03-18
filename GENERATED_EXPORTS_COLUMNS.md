# Generated Excel Files and HTML Reports: Columns, Logic, and Data Flow

Last Updated: 2026-02-20

This document describes:
- What each generated Excel column means.
- How values are derived.
- How each field flows into downstream sheets and HTML reports.

## Pipeline Order

`run_all.py` executes:

1. `run_all_exports.py`
1. `generate_nested_view_html.py`
1. `generate_missed_entries_html.py`
1. `generate_assignee_hours_report.py`
1. `generate_gantt_chart_html.py`
1. `generate_phase_rmi_gantt_html.py`
1. `fetch_jira_dashboard.py`
1. `export_ipp_phase_breakdown.py`
1. `generate_ipp_meeting_dashboard.py`

Inside `run_all_exports.py`:

1. `export_jira_subtask_worklogs.py` -> `2_jira_subtask_worklogs.xlsx`  
   **Note:** In the canonical/colossal-fetch setup, worklogs are produced by the compatibility bridge from `canonical_worklogs`; this script is not used.
1. `export_jira_work_items.py` -> `1_jira_work_items_export.xlsx`
1. `export_jira_subtask_worklog_rollup.py` -> `3_jira_subtask_worklog_rollup.xlsx`
1. `export_jira_nested_view.py` -> `nested view.xlsx`

## Shared Start Date Resolver

All scripts that need Jira planned start date now use shared resolver in `ipp_meeting_utils.py` (`resolve_jira_start_date_field_id`).

Resolver behavior:
1. Load all Jira fields.
1. Keep candidates whose name contains `start date`.
1. If multiple candidates exist, pick the one with highest non-empty coverage across in-scope issue types (`Epic`, `Story`, `Task`, `Sub-task`, `Subtask`, `Bug Task`, `Bug Subtask`) for selected projects.
1. Tie-break fallback: exact `Start date`, then first matching `start date` name.

## Shared IPP Flags

### `Latest IPP Meeting`
- Source: `IPP_MEETING_XLSX_PATH` workbook.
- Logic: Jira keys extracted from workbook text; row is `Yes` if issue key exists.

### `Jira IPP RMI Dates Altered`
- Compares IPP planned dates vs Jira Epic planned dates.
- Jira epic planned dates are fetched using shared start-date resolver + `duedate`.
- `Yes` if start/end differ; else `No`.

### `IPP Actual Date (Production Date)` + `IPP Remarks`
- Source: `IPP_MEETING_XLSX_PATH` workbook, parsed across all sheets.
- Key resolution: Jira/Epic-link columns; issue keys are normalized.
- If multiple rows exist per Jira key, the row with latest parsed IPP actual date is selected.
- `IPP Remarks` follows the selected row.

### `IPP Actual Date Matches Jira End Date`
- Compares selected IPP actual date vs Jira Epic `planned_end` (`duedate`).
- `Yes` only when both dates exist and are identical (`YYYY-MM-DD`).

## 1) `1_jira_work_items_export.xlsx`

One row per Jira issue in configured projects.

| Column | Logic | Data Flow to Other Artifacts |
|---|---|---|
| `project_key` | Jira `project.key`. | Used by `fetch_jira_dashboard.py` for project filters and grouping in `dashboard.html`; used by `generate_gantt_chart_html.py` for matching rows with issue hierarchy. |
| `issue_key` | Jira issue key. | Primary ID for dashboard epics/stories/subtasks/bug-subtasks and gantt actual-date lookup. |
| `work_item_id` | Key extracted from issue URL (fallback: key). | Export-only identifier; not a required downstream join key. |
| `work_item_type` | Normalized type. `Bug Subtask` is preserved as its own type. | Informational in file 1; downstream classification uses `jira_issue_type`. |
| `jira_issue_type` | Raw Jira issue type. | `fetch_jira_dashboard.py` uses this to route rows into `epics`, `stories`, `subtasks`, `bug_subtasks`. |
| `summary` | Jira summary. | Displayed in `dashboard.html`; used in gantt row matching keys. |
| `status` | Jira status name. | Displayed and color-coded in `dashboard.html`. |
| `start_date` | Dynamic Jira start-date field value (shared resolver selected field). | Becomes `jira_start_date` in dashboard data; influences date filtering and cards in `dashboard.html`. |
| `end_date` | Jira `duedate`. | Becomes `jira_end_date` in dashboard; date filtering and display. |
| `actual_start_date` | Earliest worklog timestamp rolled up by issue scope. | Used by dashboard card Actual timeline and by `generate_gantt_chart_html.py` for actual bars. |
| `actual_end_date` | Latest worklog timestamp rolled up by issue scope. | Used by dashboard card Actual timeline and gantt actual bars. |
| `original_estimate` | Jira estimate text. | Used by dashboard man-days rendering. |
| `original_estimate_hours` | `timeoriginalestimate` seconds -> hours. | Export visibility; not directly consumed by current HTML generators. |
| `assignee` | Jira assignee or `Unassigned`. | Used in dashboard assignee filters and card metadata. |
| `total_hours_logged` | `aggregatetimespent`/`timespent` seconds -> hours. | Used in dashboard metrics (`total_hours_logged`, epic subtask aggregation baseline). |
| `priority` | Jira priority. | Export-only currently. |
| `parent_issue_key` | Parent from `parent`/epic-link logic. | Used by dashboard to derive story->epic and subtask->story relationships. |
| `parent_work_item_id` | Parent key extracted from parent URL. | Export-only currently. |
| `parent_jira_url` | Parent browse URL. | Export-only currently. |
| `jira_url` | Issue browse URL. | Display links in dashboard cards. |
| `Latest IPP Meeting` | Shared IPP logic. | Displayed as badge in dashboard cards; merged with rows from files 2/3 at epic level. |
| `Jira IPP RMI Dates Altered` | Shared IPP-vs-Jira epic-date compare. | Displayed as sync badge in dashboard cards; merged with rows from files 2/3 at epic level. |
| `IPP Actual Date (Production Date)` | Shared IPP actual-date loader (latest row per key). | Flows to dashboard item payload; used for epic-level merge and note gating. |
| `IPP Remarks` | Shared IPP actual-date/remarks loader. | Flows to dashboard cards + markdown notes when mismatch and date-match condition pass. |
| `IPP Actual Date Matches Jira End Date` | Shared IPP-vs-Jira end-date compare. | Enables conditional green IPP note in dashboard alerts and markdown notes. |
| `created` | Jira created timestamp. | Export-only currently. |
| `updated` | Jira updated timestamp. | Export-only currently. |

## 2) `2_jira_subtask_worklogs.xlsx`

One row per subtask worklog entry.

| Column | Logic | Data Flow to Other Artifacts |
|---|---|---|
| `issue_link` | Subtask browse URL. | Used for row carry-forward into file 3 only. |
| `issue_id` | Subtask key. | Group key in file 3; used for dashboard start-date backfill from file 3. |
| `issue_title` | Subtask summary. | Carry-forward into file 3 only. |
| `issue_type` | Raw subtask-type text. | Informational in file 3. |
| `parent_story_link` | Parent story browse URL. | Carry-forward into file 3. |
| `parent_story_id` | Parent story key. | Used by file 3 and dashboard for story linkage. |
| `parent_epic_id` | Epic resolved from parent story. | Used for epic flag merges in dashboard. |
| `issue_assignee` | Subtask assignee text. | Carry-forward into file 3 only. |
| `Latest IPP Meeting` | Shared IPP logic. | Merged into epic flags in dashboard flow. |
| `Jira IPP RMI Dates Altered` | Shared IPP logic. | Merged into epic flags in dashboard flow. |
| `IPP Actual Date (Production Date)` | Shared IPP actual-date loader by parent epic. | Merged at epic level in dashboard flow (latest date kept). |
| `IPP Remarks` | Shared IPP remarks by parent epic. | Merged at epic level in dashboard flow (first non-empty kept). |
| `IPP Actual Date Matches Jira End Date` | Shared date-equality check by parent epic. | Merged into epic yes/no note condition in dashboard flow. |
| `worklog_started` | Jira worklog started timestamp. | Rolled to min/max in file 3; also upstream source for file 1 actual date rollups. |
| `hours_logged` | Worklog seconds -> hours. | Summed in file 3; added into epic subtask hours in dashboard flow. |

## 3) `3_jira_subtask_worklog_rollup.xlsx`

One row per subtask (grouped from file 2).

| Column | Logic | Data Flow to Other Artifacts |
|---|---|---|
| `issue_link` | First non-empty from grouped rows. | Informational only. |
| `issue_id` | Subtask key group. | Used by dashboard to locate subtask/bug-subtask for planned-start backfill when row start is blank. |
| `issue_title` | First non-empty grouped value. | Informational only. |
| `issue_type` | First non-empty grouped value. | Informational only. |
| `parent_story_link` | First non-empty grouped value. | Informational only. |
| `parent_story_id` | First non-empty grouped value. | Used by dashboard to update related story planned start. |
| `parent_epic_id` | First non-empty grouped value. | Used by dashboard to update related epic planned start and merge flags. |
| `issue_assignee` | First non-empty grouped value. | Informational only. |
| `Latest IPP Meeting` | Shared IPP logic. | Merged into epic-level dashboard flag. |
| `Jira IPP RMI Dates Altered` | Shared IPP logic. | Merged into epic-level dashboard sync status. |
| `IPP Actual Date (Production Date)` | Shared IPP actual-date loader by parent epic. | Merged into epic-level dashboard field (latest date wins). |
| `IPP Remarks` | Shared IPP remarks by parent epic. | Merged into epic-level dashboard notes source. |
| `IPP Actual Date Matches Jira End Date` | Shared date-equality check by parent epic. | Merged into epic-level condition for remark-note display. |
| `planned start date` | Parent story Jira planned start (shared resolver field). | Dashboard may apply as earliest-date merge for epic/story and fallback for subtask/bug-subtask when own start missing. |
| `planned end date` | Parent story Jira `duedate`. | Stored in file 3; not currently used by dashboard merge logic. |
| `actual start date` | Min `worklog_started` for subtask. | Informational in file 3; detailed actuals primarily come from file 1 in dashboard and gantt. |
| `actual end date` | Max `worklog_started` for subtask. | Informational in file 3. |
| `total hours_logged` | Sum of grouped `hours_logged`. | Consumed by `export_jira_nested_view.py` for subtask `Actual Hours` in `nested view.xlsx`. |

## 4) `nested view.xlsx`

Hierarchical sheet used by HTML reports.

| Column | Logic | Data Flow to Other Artifacts |
|---|---|---|
| `Aspect` | Row label by hierarchy level. | Rendered directly in `nested_view_report.html`; used as category labels in `gantt_chart_report.html`. |
| `Man-days` | `Man-hours / 8`. | Rendered in nested-view report and gantt metadata table columns. |
| `Man-hours` | Jira estimate seconds -> hours. | Rendered in nested-view report and gantt metadata. |
| `Actual Hours` | Subtasks use file 3 rollup; higher levels use Jira aggregates. | Rendered in nested-view report and gantt metadata. |
| `Actual Days` | `Actual Hours / 8`. | Rendered in nested-view report and gantt metadata. |
| `Planned Start Date` | Jira start date (shared resolver), with subtask fallback from file 3 when available. | Used for planned bars in `gantt_chart_report.html`; rendered in nested-view report tables. |
| `Planned End Date` | Jira `duedate` (with subtask fallback from file 3 when available). | Used for planned bars in `gantt_chart_report.html`; rendered in nested-view report tables. |

## 5) `assignee_hours_report.xlsx`

Normalized one-row-per-worklog summary used by the assignee-hours HTML report.

| Column | Logic | Data Flow to Other Artifacts |
|---|---|---|
| `project_key` | Derived from `issue_id` prefix before `-` in file 2. Fallback: `UNKNOWN`. | Used by `assignee_hours_report.html` project multi-select filter. |
| `worklog_date` | Date parsed from file 2 `worklog_started` timestamp (`YYYY-MM-DD`). | Primary date-range filter input in `assignee_hours_report.html`. |
| `period_day` | Same as `worklog_date`. | Used when granularity is `day`. |
| `period_week` | ISO week code from `worklog_date` (`YYYY-Www`, Monday-based). | Used when granularity is `week`. |
| `period_month` | Year-month from `worklog_date` (`YYYY-MM`). | Used when granularity is `month`. |
| `issue_assignee` | File 2 `issue_assignee`, defaulting to `Unassigned` when empty. | Group-by key in `assignee_hours_report.html`. |
| `hours_logged` | File 2 `hours_logged` numeric value (positive rows only). | Summed by period + assignee in `assignee_hours_report.html`. |

## Dashboard Data Contract (`fetch_jira_dashboard.py` -> `dashboard.html`)

### Top-level arrays
- `epics`: file 1 rows where issue type contains `epic`.
- `stories`: file 1 rows where issue type is `story`.
- `subtasks`: file 1 rows where issue type is `sub-task`/`subtask`.
- `bug_subtasks`: file 1 rows where issue type is `bug subtask`.

### Orphans
- `orphans.stories`: story rows whose `epic_key` not found in `epics`.
- `orphans.subtasks`: subtask rows whose `story_key` not found in `stories`.
- `orphans.bug_subtasks`: bug-subtask rows whose `story_key` not found in `stories`.

### Item fields (common shape across all arrays)
| Field | Source Logic | Data Flow in `dashboard.html` |
|---|---|---|
| `issue_key` | File 1 `issue_key`. | Card identity, copy/export, relation filters. |
| `issue_type` | Normalized label in dashboard builder. | Card type label. |
| `summary` | File 1 `summary`. | Card title, search matching. |
| `assignee` | File 1 `assignee`. | Lane filters and card chips. |
| `project_key` | File 1 `project_key` (fallback from key prefix). | Project chip filtering. |
| `jira_url` | File 1 `jira_url` fallback constructed link. | Card Jira link. |
| `parent_issue_key` | File 1 parent key. | Relationship derivation. |
| `epic_key` | Derived by type + parent linkage. | Story/subtask/bug-subtask scoping under selected epic. |
| `story_key` | Derived for subtask/bug-subtask from parent. | Child scoping under selected story. |
| `jira_start_date` | File 1 start date, then earliest-date merge with file 3 planned start where applicable. | Date chips, range filter, timeline display. |
| `jira_end_date` | File 1 end date (planned end aliases supported). | Date chips, range filter. |
| `actual_start_date` | File 1 `actual_start_date`. | Actual timeline display on cards. |
| `actual_end_date` | File 1 `actual_end_date`. | Actual timeline display on cards. |
| `original_estimate` | File 1 estimate text. | Planned man-days conversion in cards. |
| `total_hours_logged` | File 1 total hours. | Card metrics and epic totals. |
| `latest_ipp_meeting` | File 1 + merged Yes/No from files 2/3 at epic level. | Badge rendering and markdown export. |
| `jira_ipp_rmi_dates_altered` | File 1 + merged Yes/No from files 2/3 at epic level. | Sync status badge and alerts. |
| `ipp_planned_start_date` | IPP workbook planned-start date by epic key (`load_ipp_planned_dates_by_key`), propagated to epic/story/subtask/bug-subtask by epic linkage. | Rendered in red on cards only when mismatch is `Yes`; included in copied markdown notes. |
| `ipp_planned_end_date` | IPP workbook planned-end date by epic key (`load_ipp_planned_dates_by_key`), propagated to epic/story/subtask/bug-subtask by epic linkage. | Rendered in red on cards only when mismatch is `Yes`; included in copied markdown notes. |
| `ipp_actual_date` | File 1 + merged latest-date candidate from files 2/3 at epic level. | Supports audit/debug context for IPP actual-date alignment. |
| `ipp_remarks` | File 1 + first non-empty merged remarks from files 2/3 at epic level. | Rendered in dashboard green note + included in copied markdown notes. |
| `ipp_actual_matches_jira_end_date` | File 1 + merged Yes/No from files 2/3 at epic level. | Gates green IPP note rendering when mismatch alert is active. |
| `status` | File 1 status. | Card status, color class, search. |

## HTML Report Inputs Summary

### `nested_view_report.html`
- Source: `nested view.xlsx`.
- Uses all sheet columns directly for hierarchy table rendering.
- Header date-range filter is applied to work rows (`Epic`, `Story`, `Subtask`) using Jira planned dates with OR logic:
  - include when `Planned Start Date` is in selected month range OR `Planned End Date` is in selected month range.
- Default month window is fixed to `2026-01` through `2026-02` (January 2026 to February 2026).
- Adds derived columns at HTML-generation time:
  - `Delta Days = Man-days - Actual Days` (planned minus logged)
  - `Delta Hours = Man-hours - Actual Hours` (planned minus logged)
  - `Resource Logged` = `Yes` when `Actual Hours > 0`, else `No` (for work rows: `Epic`, `Story`, `Subtask`; blank for non-work rows).
- Metrics display order in the table:
  - `Type, Assignee, Man-days, Actual Days, Delta Days, Man-hours, Actual Hours, Delta Hours, Resource Logged`
- Table viewport behavior:
  - `Aspect` column is sticky during horizontal scrolling.
  - `Aspect` column width can be expanded/collapsed by dragging the right border of the `Aspect` header; preference is saved in browser local storage.
  - Desktop table uses full available panel width to display as many columns as fit by viewport size.
  - Remaining columns are available via horizontal scroll to the right.
- Delta color semantics:
  - negative: red
  - positive: green
  - zero: white
- Tree controls:
  - `Expand All`
  - `Collapse To Projects` (collapses at project level and hides `Category` rows by switching product categorization off)
  - `Collapse to Epics` (keeps Project/Category/Epic visible and collapses Epic descendants)
- `No Entry <N>` control: `N` is count of work rows (`Epic`, `Story`, `Subtask`) where any of `Man-days`, `Man-hours`, or `Actual Hours` is `0`.
- When `N = 0`, the `No Entry <N>` control is disabled.
- When `N > 0`, the `No Entry <N>` control is shown in red and can toggle filtered view.
- Assignee rendering:
  - `Subtask`: direct assignee(s) from child `Assignee` rows.
  - `Story`: aggregated assignee set from descendant subtasks.
  - `Epic`: aggregated assignee set from descendant subtasks.

### `missed_entries.html`
- Source: `1_jira_work_items_export.xlsx`.
- Scope: all work-item rows present in file 1 (`Epic`, `Story`, `Sub-task`, `Bug Subtask`).
- Date filter:
  - month range controls (`From`, `To`, `Reset`)
  - default month window is previous month through current month
  - row included when selected month range overlaps available planned dates (`start_date` and/or `end_date`)
  - rows with both planned dates missing are excluded from date-filtered results.
- Missing-field selector (multi-select, ANY semantics):
  - `Jira Planned Start Date` (`start_date`)
  - `Jira Planned Due Date` (`end_date`)
  - `Original Estimates` (`original_estimate`)
- Outputs:
  - assignee summary table with total and per-field missing counts
  - assignee name acts as accordion toggle to show a per-work-item detail table
  - detail table columns include `Work Item`, `Issue Type`, `Missing Fields`, and `Resource Logged Hours`
  - `Resource Logged Hours` is `Yes` when file 1 `total_hours_logged` is greater than `0`, otherwise `No`.

### `assignee_hours_report.html`
- Source: `assignee_hours_report.xlsx` (sheet: `AssigneeHours`).
- Controls:
  - granularity dropdown: `day`, `week`, `month`
  - date range controls: `From`, `To`, `Reset` (`type="date"`)
  - project multi-select (`project_key`)
- Default date range: previous month start through current month end, bounded by available data when needed.
- Grouping and totals:
  - grouped by selected period (`period_day`/`period_week`/`period_month`) and `issue_assignee`
  - totals shown as summed `hours_logged` per grouped row
  - ISO weekly grouping (Monday-start weeks)
- Output table columns:
  - `Period`
  - `Assignee`
  - `Total Hours`

### `gantt_chart_report.html`
- Sources: `nested view.xlsx` + `1_jira_work_items_export.xlsx`.
- Planned bars: `Planned Start Date` and `Planned End Date` from `nested view.xlsx`.
- Actual bars: matched from file 1 `actual_start_date`/`actual_end_date` by hierarchy keys.

### `phase_rmi_gantt_report.html`
- Source: `nested view.xlsx` (Jira-derived planned dates/man-days).
- Treats `Epic` rows as RMIs and child `Story` rows as phases.
- Fixed phase lanes (in order): `Research/URS`, `DDS`, `Development`, `SQA`, `User Manual`, `Production`.
- Aggregates phase workload by `(phase_name, rmi_name)`:
  - `planned_start = min(story planned start)`
  - `planned_end = max(story planned end)`
  - `man_days = sum(story man_days)`
- UI:
  - date range selector only (`From`, `To`, `Reset`)
  - default window = previous month start through next month end (3-month span centered on current month)
  - sticky phase lane column and sticky weekly/month header
  - horizontal scroll timeline
  - one dedicated lane per fixed phase
  - mini RMI cards show `RMI Name`, `Man Days`, `Planned Start Date`, `Planned End Date`
  - lane-level weekly load chips indicate overload intensity by week.

### `dashboard.html`
- Source: JSON payload produced by `fetch_jira_dashboard.py` from files 1, 2, 3.
- Uses `epics`, `stories`, `subtasks`, `bug_subtasks`, plus `orphans` groups.
- Bug Subtasks are rendered as story children and sibling lane group to subtasks, using same filtering and calculation logic.
- Includes a global `remove filters` control that resets projects, search text, assignee filters, status filters, date range, and lane-card selections (`epic`/`story`) in one action.
- Global filter reset sets date bounds to the full dataset month span (earliest to latest month across dashboard items); if no dated items exist, defaults are used.
- When no epic/story card is selected, lane scoping is disabled: stories, subtasks, and bug subtasks show all items that match active lane filters.
- Date-mismatch alert remains red; when mismatch is `Yes`, card date tables show an additional red `IPP Meeting` row with IPP planned dates and copied markdown adds `IPP planned dates: <start> to <end>`.
- When mismatch is `Yes` and `ipp_actual_matches_jira_end_date` is `Yes` and `ipp_remarks` is non-empty, a green IPP note is shown in card UI and markdown export.

## Environment Variables Affecting Flows

- `JIRA_PROJECT_KEYS`: project scope used by exports and start-date dominance sampling.
- `IPP_MEETING_XLSX_PATH`: source workbook for both IPP flags.
- `JIRA_EXPORT_XLSX_PATH`: file 1 output/input path.
- `JIRA_WORKLOG_XLSX_PATH`: file 2 output/input path.
- `JIRA_SUBTASK_WORKLOG_INPUT_XLSX_PATH`: file 3 input path.
- `JIRA_SUBTASK_ROLLUP_XLSX_PATH`: file 3 output/input path.
- `JIRA_NESTED_VIEW_XLSX_PATH`: file 4 output/input path.
- `JIRA_MISSED_ENTRIES_HTML_PATH`: output HTML path for `generate_missed_entries_html.py` (default: `missed_entries.html`).
- `JIRA_ASSIGNEE_HOURS_XLSX_PATH`: output XLSX path for `generate_assignee_hours_report.py` (default: `assignee_hours_report.xlsx`).
- `JIRA_ASSIGNEE_HOURS_HTML_PATH`: output HTML path for `generate_assignee_hours_report.py` (default: `assignee_hours_report.html`).
- `JIRA_PHASE_GANTT_INPUT_XLSX_PATH`: source workbook path for `generate_phase_rmi_gantt_html.py` (default: `nested view.xlsx`).
- `JIRA_PHASE_GANTT_HTML_PATH`: output HTML path for `generate_phase_rmi_gantt_html.py` (default: `phase_rmi_gantt_report.html`).
- `JIRA_PRODUCT_CATEGORIZATION_FIELD_ID`: optional field override for nested view product grouping.

## Run Commands

- Full refresh (exports + reports): `python run_all.py`
- Server only (no generation): `python run_server.py`
- Exports only: `python run_all_exports.py`
- Dashboard only: `python fetch_jira_dashboard.py`
- Nested view HTML only: `python generate_nested_view_html.py`
- Missed entries HTML only: `python generate_missed_entries_html.py`
- Assignee-hours HTML/XLSX only: `python generate_assignee_hours_report.py`
- Gantt HTML only: `python generate_gantt_chart_html.py`
- Phase-owner RMI gantt HTML only: `python generate_phase_rmi_gantt_html.py`
