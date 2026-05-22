# Monthly Epic Plan Progress Report

## Purpose

The Monthly Epic Plan Progress report compares the epics in scope for a selected month against Jira execution data. It uses the report's existing epic scope logic: epics whose approved start or due date falls in the selected month, plus unresolved brought-forward epics inside the configured overdue lookback window.

## User Flow

1. Open `monthly_epic_plan_progress_report.html`.
2. Select the month, project scope, employee scope, effort unit, epic scope mode, overdue lookback, and date-filter toggles.
3. Review the executive summary, estimate hierarchy stats, employee stats, resource planning, project cards, and epic table or Gantt view.

## Fields And Validations

| Area | Field | Type | Behavior |
|---|---|---|---|
| Controls | Month | Month input | Drives the selected calendar month for epic and worklog scope. |
| Controls | Projects | Multi-select dropdown | Restricts the API payload to selected project keys. |
| Controls | Employees | Hierarchical checkbox dropdown | Restricts workforce capacity and leave calculations. |
| Controls | Effort unit | Radio buttons | Switches displayed effort values between hours and days. |
| Controls | Epic scope | Segmented buttons | Chooses TK planner epics, all epics, or all Jira epics mode. |
| Controls | Overdue lookback | Numeric input | Controls how many days before the selected month can be brought forward. |
| Estimate hierarchy stats | Month Plan | Bar + calculated number | Description: Executive planned hours for epics planned this month. This equals the Executive summary planned-hours value after excluding brought-forward overdue epics. |
| Estimate hierarchy stats | Epic Estimate | Bar + calculated number | Description: Sum of Epic Original Estimates. Sums direct Jira epic `original_estimate_hours` for epics planned in the selected month. |
| Estimate hierarchy stats | Story Estimate | Bar + calculated number | Description: Sum of their Epics' Stories Original Estimate. Sums direct story `original_estimate_hours` for stories planned in the selected month or parent stories of subtasks planned in the selected month. |
| Estimate hierarchy stats | Subtask Estimate | Bar + calculated number | Description: Sum of Story's subtasks Original Estimates. Sums subtask `original_estimate_hours` for subtasks planned in the selected month. |
| Estimate hierarchy stats | Subtask Logged | Bar + calculated number | Description: Sum of Story's subtasks Logged Hours. Sums selected-month logged hours on subtasks planned in the selected month. |
| Estimate hierarchy stats | Story Overrun | Bar + calculated number and percent | Description: Logged hours exceeding parent story original estimates. Sums per-story selected-month overrun where subtask logged hours exceed that parent story's original estimate. |
| Estimate hierarchy detail drawer | Work item table | Right-side resizable drawer | Opens when any estimate bar is clicked. Shows Jira link, work item name, parent links for lower hierarchy items, planned dates, original estimates, planner-backed TK planned hours for matching epic/story records where available, logged hours where available, and overrun where applicable. |

## Business Rules

- The estimate hierarchy stats use the report's planned-this-month epic subset. Brought-forward overdue epics remain visible in the Executive summary and epic table but do not inflate these month-plan hierarchy totals.
- The estimate hierarchy stats still honor project, month, scope mode, on-hold handling, and client-side start/end/range toggles.
- Month Plan is a reference bar that matches the Executive summary planned-hours number for the same visible month-planned epic set.
- Direct epic estimates, direct story estimates, and direct subtask estimates are shown separately. The hierarchy stats do not replace the report's existing planned-hours calculation.
- Subtask logged overrun is calculated per parent story for the selected month: `max(sum(month subtask logged hours) - story original estimate hours, 0)`.
- Clicking an estimate bar opens a detail drawer scoped to the currently visible report rows. Story and subtask rows include parent Jira links so users can trace the hierarchy.
- The estimate detail drawer adds a `TK planned` column sourced from the epics-planner database when the work item matches a planner-backed epic or a planner phase Jira link. Rows without a matching planner record keep this value blank.
- The estimate detail drawer width can be adjusted by dragging the handle on the drawer's left edge.
- Overrun percentage uses only the parent-story estimate baseline for stories that were exceeded.
- Days are derived from the report metadata `hours_per_day`, currently `8`.

## Dependencies

- `monthly_epic_plan_progress_service.py` builds the API payload, estimate hierarchy rollups, planner-to-Jira TK planned-hour mappings, and per-metric work item detail rows.
- `monthly_epic_plan_progress_report.html` renders the report, client-side date toggles, estimate hierarchy bar chart, resizable estimate detail drawer, table, and Gantt.
- `report_server.py` serves `/api/monthly-epic-plan-progress/summary`.
- SQLite tables: `canonical_issues`, `canonical_worklogs`, `canonical_refresh_state`, and `epics_management`.

## Related Code

- `monthly_epic_plan_progress_service.py`
- `monthly_epic_plan_progress_report.html`
- `report_server.py`
- `tests/test_monthly_epic_plan_progress.py`

## Change Notes

- Added the Estimate hierarchy stats section with direct epic, story, subtask estimate, subtask logged, and parent story overrun metrics.
- Updated the Estimate hierarchy stats UI from long-label cards to short-name horizontal bars with descriptions.
- Added clickable estimate bars that open a resizable right-side drawer with the work items considered for the selected stat.
- Added planner-backed `TK planned` values to estimate-detail drawer rows for epic and story issues when the epics planner has a matching Jira-linked record.
