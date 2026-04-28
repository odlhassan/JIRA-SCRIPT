# RMI Jira Gantt Report

Report ID: `rmi_jira_gantt`

Generator: `generate_rmi_jira_gantt_html.py`
Output: `rmi_jira_gantt_report.html` (root + `report_html/`)
Served at: `http://127.0.0.1:3000/rmi_jira_gantt_report.html`

## Purpose

Single-file HTML dashboard combining the Epics Planner workbook plan with the
canonical Jira tree, showing RMI epic timelines, hierarchical breakdowns,
capacity vs TK Approved, and an executive summary card row at the top of the
page.

### Capacity — employee breakdown drawer

Click **No. of Employees** in the capacity calculator to open the same
right-hand drawer used for metric drill-downs. It lists who is **considered**
for the current headcount and who is **not**, with short reasons (depends on
mode: org FTE profile, team multi-select, or product filter). The drawer reuses
`#detail-drawer` and explains alignment between performance teams and the RLT
assignee list when using the policy FTE headcount.

## Executive Summary Cards

The card row at the top of the page is rendered by `renderMetrics()` and
`renderProductSummary()` in `generate_rmi_jira_gantt_html.py`. All values are
computed live from `scopedTotals()` over the currently filtered epic set; no
values are hardcoded.

| Card | Label | Source field (per epic) | Aggregation |
| --- | --- | --- | --- |
| Total # of RMI Epics | `Epic parents in the selected product scope` | epic row | `count(scoped epics)` |
| Optimistic | `Workbook optimistic total for the selected epic set` | `optimistic_seconds` | `sum` |
| Most Likely | `Workbook most likely total for the selected epic set` | `most_likely_seconds` | `sum` |
| Pessimistic | `Workbook pessimistic total for the selected epic set` | `pessimistic_seconds` | `sum` |
| Calculated | `Workbook calculated estimate total for the selected epic set` | `calculated_seconds` | `sum` |
| TK Approved (hero) | `TK approved total for the selected epic set` | `tk_approved_seconds` | `sum` |
| Idle Hours/Days | `Total Availability minus TK Approved for the current scope.` (or `TK Approved exceeds total availability` when negative) | `availability − tk_approved_seconds` | derived |
| Epic Estimates | `Epic-level Jira original estimate total` | `jira_original_estimate_seconds` | `sum` |
| Story Estimates | `Story-level original estimate total (excludes subtasks)` | `story_estimate_seconds` | `sum` |
| Subtask Estimates | `Subtask-level original estimate total` | `subtask_estimate_seconds` | `sum` |
| Logged | `Total hours logged across all stories/subtasks` | `logged_seconds` | `sum` |

### Product Summary Cards

Below the metric grid, one summary card renders per product (plus an `All
Products` aggregate). Accent colors:

| Product | Accent |
| --- | --- |
| Digital Log | `#7c3aed` |
| Fintech Fuel | `#b45309` |
| OmniChat | `#2563eb` |
| OmniConnect | `#0f766e` |
| All Products | `#102033` |

Each card shows `Total TK Approved` (`sum(tk_approved_seconds)`) and the count
of RMIs/Epics in that product.

## Data Sources

All fields resolve through `load_report_data(db_path, run_id)` which reads from
the canonical `assignee_hours_capacity.db`:

- Canonical run id resolved via `resolve_canonical_run_id` (in
  `canonical_report_data.py`).
- Epic / story / subtask tree and worklogs come from the canonical Jira
  snapshot tables.
- `optimistic_seconds`, `most_likely_seconds`, `pessimistic_seconds`,
  `calculated_seconds`, `tk_approved_seconds` come from the
  `epics_management` planner workbook embedded in the canonical DB
  (`_epic_metrics`).
- **Capacity / leaves / headcount (capacity calculator block)** match the
  [Nested view report](01-nested-view-report.md) scorecard: **planned leave
  hours** are aggregated from the same RLT leave workbook
  (`rlt_leave_report.xlsx` by default, or `JIRA_LEAVE_REPORT_XLSX_PATH`):
  `Daily_Assignee` rows with `planned_taken_hours` + `planned_not_taken_hours`
  per `period_day`, bucketed to `YYYY-MM` (same as nested view
  `totalPlannedLeavesHours`). **Employee count** for the “all products / all
  teams” scope uses the overlapping row from `assignee_capacity_settings` in
  `assignee_hours_capacity.db` (via `_load_capacity_profiles` — the same
  profile table nested view uses), not the count of RMI-scoped Jira
  assignees. When you filter by product or team, leave hours are the sum of
  matching assignee rows; headcount comes from the scoped assignee list.
- **Teams** for the capacity team selector come from `performance_teams` in
  the canonical DB. `rlt_employees` from the RLT Jira snapshot is still
  available for other filters but is not the source of month-level total
  planned leave in global scope.
- The capacity selector is a two-level dropdown: **Teams** with nested
  **Members**. Selection is member-centric (deduped globally), supports
  **Select all**, **Clear selection**, team-level toggle, and member-level
  toggle. Parent team checkboxes show an indeterminate state when only some
  members are selected.

## Related Code

- `generate_rmi_jira_gantt_html.py` — generator and embedded JS for cards.
- `canonical_report_data.py` — `resolve_canonical_run_id`, planner fetch.
- `assignee_hours_capacity.db` — canonical source of truth (Jira tree,
  planner, capacity).

## Change Notes

- 2026-04-27 (later): Aligned capacity calculator with nested view: monthly
  planned leave totals from `rlt_leave_report.xlsx` (not RLT Jira
  `build_rlt_leave_snapshot` month aggregates) and org headcount from
  `assignee_capacity_settings` for the month.
- 2026-04-28: Upgraded Capacity Calculator Team filter to a two-level
  Team→Members dropdown with member-level selection, team indeterminate states,
  and bulk controls (Select all / Clear selection).
- 2026-04-27: Restored full executive summary card row matching the
  `IPP Meeting Reports/rmi_jira_gantt.html` reference (Total # of RMI Epics,
  Optimistic / Most Likely / Pessimistic / Calculated, TK Approved hero,
  Idle Hours/Days, Epic / Story / Subtask Estimates, Logged) plus per-product
  summary cards. Reference accent palette applied. All values driven by
  canonical DB; nothing hardcoded.
