# Epic Explorer Report

Report ID: `epic_explorer`

INFO_IDS:  

## Purpose

Epic Explorer is a canonical-database report for inspecting every Jira epic and drilling from epic to work items, subtasks, bug subtasks, and worklogs. It is designed for delivery review, estimate reconciliation, and epic-level execution analytics.

## User Flow

1. Open `epic_explorer_report.html`.
2. Leave filters blank to load all canonical Jira epics.
3. Optionally apply a From/To date range or project selection. Date filters and project filters have separate Apply/Clear controls. These filters only include or exclude epic rows; they do not trim the nested work items, subtasks, bug subtasks, worklogs, or drawer analytics inside each included epic.
4. Expand an epic row to see work items, expand a work item to see subtasks and bug subtasks, then expand a subtask to see individual worklogs.
5. Click the epic name to open the right-side analytics drawer. Drag the drawer's left edge to resize from half page width toward full width.
6. Click the Jira open icon beside an epic name to open the epic in Jira.
7. Click CSV to export the visible top-level epic table.
8. Use the Executive Summary section above the main table to pin one or more epics (via the "Add epics" dropdown with checkboxes and a search box) into a leadership-ready mini dashboard. Typing in the search box filters the checkbox list without closing or replacing the focused search field. Each checkbox change immediately refreshes the pinned-epic count and charts, and Apply simply closes the picker after saving the live selection; pinned epics persist across reloads via browser local storage.
9. In the Executive Summary mini dashboard, click an epic name to open the same detailed analytics drawer used by the main table. Click the chevron icon to quick-expand a week-over-week schedule variance trend for that epic without leaving the mini dashboard; the expanded panel now includes a line chart built from the same weekly values shown in the table. Click the close icon to unpin an epic and refresh the dashboard immediately.
10. Review the Month Over Month Average Schedule Variance Trend chart and the Portfolio Budget vs Actual Hours chart below the mini dashboard table for a portfolio-level view of the pinned epics.
11. Click an SV Hours value in the main epic table to open its calculation breakdown. The dialog identifies the estimate source, calendar period, reporting date, elapsed days, planned-to-date proration, actual hours, signed variance, and variance percentage.

## Fields And Validations

| Area | Field | Type | Default | Behavior |
|---|---|---|---|---|
| Filters | From | Date input | Blank | When both dates are set, includes only epics whose Jira epic `start_date` to `due_date` overlaps the range. |
| Filters | To | Date input | Blank | Pairs with From for overlap filtering. If only one side reaches the API, the server treats it as a single-day filter. |
| Actions | Apply dates / Clear dates | Buttons | - | Apply or clear only the date overlap filter. Project selections remain as-is. |
| Filters | Projects | Checkbox menu | All projects | Shows project display names when Epics Planner metadata is available. Includes only epics whose canonical epic `project_key` matches selected projects after Apply projects is clicked. |
| Actions | Apply projects / Clear projects | Buttons | - | Apply or clear only the project filter. Date filters remain as-is. |
| Filters | Capacity basis | Select | Assignee capacity after leaves | Saves the monthly plan allocation basis to the database and reloads the report. `Assignee capacity after leaves` weights story estimates by configured workday hours, Ramadan hours, holidays, and assignee leave. `Standard workdays` splits by normal workdays without leave adjustment. |
| Actions | CSV | Button | - | Exports visible top-level epic rows to `epic_explorer.csv`. |
| Table | # | Sticky numeric column | Generated | Shows the visible row number for easier scanning. |
| Table | Epic Name | Button plus Jira icon | Derived | Name opens the analytics drawer; icon opens Jira in a new tab. |
| Table | Assignee | Text | Derived | Canonical epic `assignee`. |
| Table | Product | Text | Blank when missing | `epics_management.product_category` matched by epic key. |
| Table | TK Budget | Numeric hours | Blank when missing | Epic-plan `tk_budgeted_man_days` or `man_days` multiplied by 8. |
| Table | Jira Original Estimate | Numeric hours | 0 | Canonical epic `original_estimate_hours`. |
| Table | Story Estimates | Numeric hours | 0 | Sum of canonical parent work-item original estimates under the epic. |
| Table | Subtask Estimates | Numeric hours | 0 | Sum of canonical `Sub-task` and `Bug Subtask` original estimates under the epic. |
| Table | Jira Epics' Planned Dates | Date range | Blank when missing | Canonical epic `start_date` to `due_date`. |
| Table | Total Actual Hours | Numeric hours | 0 | Lifetime descendant actual hours for subtasks and bug subtasks. The report uses the greater of stored worklog-row totals and the canonical Jira issue `total_hours_logged` value so actuals are not limited to the refresh scope year. |
| Table | Actual Complete Date | Date | Blank when unresolved or missing | For a currently resolved epic, the later of descendant subtask last worklog date and epic `resolved_stable_since_date`. Unresolved and reopened epics remain blank. |
| Table | Planned vs Actual Hours | Numeric comparison | 0 / 0 | Planned value uses canonical epic `original_estimate_hours`; actual value uses lifetime descendant subtask and bug-subtask actuals. TK Budget, story estimates, and subtask estimates are comparison/fallback values only. |
| Table | Planned vs Actual Delivery | Date comparison | Blank when missing | Planned value uses the canonical epic `due_date`; actual value uses the calculated Actual Complete Date. |
| Table | SV Date | Signed days | Blank when missing | Jira epic planned due date minus Actual Complete Date for resolved epics, or minus today for unresolved and reopened epics. Negative means behind. |
| Table | SV Hours | Clickable signed hours and percent | Blank when no planned basis | Actual-to-date hours minus planned-to-date hours. Planned-to-date is prorated through Actual Complete Date for resolved epics and through today for unresolved and reopened epics; it reaches the full original estimate after the planned due date. Clicking the value opens the full calculation breakdown and estimate fallback source. |
| Table | Est. Accuracy | Percent | Blank when no actuals | Jira epic original estimate divided by actual hours multiplied by 100. |
| Table | Epic Status | Status pill | Derived | Canonical epic `status`. |
| Table | Headcount | Integer | 0 | Distinct worklog authors on descendant subtasks and bug subtasks. |

## Executive Summary Mini Dashboard Fields

| Area | Field | Type | Default | Behavior |
|---|---|---|---|---|
| Picker | Add epics | Dropdown with search + checkboxes | None selected | Client-side only; lists every epic currently loaded in the page's payload. Search box filters by epic name or key. Each checkbox change saves the live selection, refreshes the count and charts, and keeps the picker state in sync. |
| Table | Epic Name | Button | - | Opens the same analytics drawer as the main table. |
| Table | Status | Status pill | Derived | Row's `epic_status`, showing whether the pinned epic is resolved, in progress, on hold, or any other Jira status present in the canonical payload. |
| Table | Budget | Numeric hours | Blank when no planned basis exists | Uses the epic's Jira original estimate (`jira_original_estimate_hours`) when it is positive; otherwise falls back to `planned_total_hours`. TK budget is not used for Executive Summary Budget. |
| Table | Actual Hours | Numeric hours | 0 | Row's `total_actual_hours`. |
| Table | Planned Start / Planned Due | Date (`MMM-DD-YY`) | Blank when missing | Row's `planned_start` / `planned_due`, displayed in `MMM-DD-YY` format for the Executive Summary table. |
| Table | Actual Complete Date | Date with hover tooltip (`MMM-DD-YY`) | Blank when missing | Row's `actual_complete_date`, displayed in `MMM-DD-YY`; hovering shows plain-English reasoning derived from `actual_complete_source` (e.g. "Later of last logged worklog date and epic resolved-stable-since date."). |
| Table | SV Date / SV Hours | Signed KPI text | Blank when missing | SV Date keeps the epic schedule date basis. SV Hours is the Executive Summary budget variance: `total_actual_hours - Budget`, where Budget uses the Jira original estimate shown in the same row. |
| Table | Quick-expand toggle | Chevron button | Collapsed | Expands an inline week-over-week schedule variance panel for that epic only, without opening the full drawer. The panel includes a story-level estimate proration line chart plus a centered weekly data table. |
| Table | Remove | Icon button | - | Unpins the epic from the mini dashboard, updates local storage, and refreshes the count and charts immediately. |
| Chart | Month Over Month Average Schedule Variance Trend | Dual-line SVG chart | - | Plots, for up to the last 6 calendar months (partial current month included), the average SV Hours and an average SV Date "day-equivalent" across all pinned epics. |
| Chart | Portfolio Budget vs Actual Hours | Horizontal bar chart | - | One bar pair (budget vs actual) per pinned epic, using the same Budget fallback as the table. Budget remains blue and actual remains green; when values overlap, the lower value is rendered above the larger value so budget remains visible when actual exceeds budget. |

## Business Rules — Executive Summary

- The Executive Summary section is entirely client-side: it reuses the same `/api/epic-explorer/summary` payload already loaded for the main table and adds no new API calls or database schema.
- Pinned epic keys are stored in the browser's `localStorage` under `epicExplorerExecSummaryEpics` so the mini dashboard persists across page reloads on the same browser/device. Keys no longer present in the current payload are simply not rendered (no error).
- The epic picker search input updates only the checkbox list while the user types, preserving focus and typed text for seamless filtering.
- Executive Summary SV Hours is intentionally tied to the visible Budget and Actual Hours columns so reviewers can reconcile it directly from the row without opening the drawer. Detailed drawer schedule-variance sections still use planned-to-date schedule proration.
- Each checkbox toggle in the epic picker commits the current pinned set immediately so the count and charts stay live while the picker is open.
- Week-over-week schedule variance uses story-level precision: each work item's own `original_estimate_hours` is linearly prorated across that story's own `start_date`..`due_date` (calendar days), then summed across all stories in the epic to get planned-to-date hours for a given week-ending date. Actual-to-date hours sum descendant subtask worklogs up to that date. The variance (actual minus planned) determines whether the epic was ahead, on track, or behind for that week.
- The quick-expanded weekly trend panel renders both a line chart and the weekly table from the same calculated planned/actual rows so the chart stays aligned with the visible data.
- The quick-expanded weekly trend chart x-axis uses `MMM-DD` labels, and the Weekly Trend table `Week Ending` column uses `MMM-DD-YYYY`.
- The weekly trend table centers the header text and cell values to make the week-ending, planned, actual, variance, and status columns easier to scan.
- The Executive Summary pinned-epics table centers text across columns (including Budget and Actual Hours) for consistent scanability.
- The Executive Summary pinned-epics table keeps the Epic Name column left-aligned while the remaining columns stay centered for readability.
- If a story is missing start/due dates, its own planned-to-date contribution is treated as 0 until due-date-only completion, keeping the trend conservative rather than guessing.
- The weekly trend range runs from the earliest story start date to the earlier of the latest story due date or today, in 7-day steps ending on the epic's actual due date.
- Month-over-month averages are computed over the pinned epics for up to the last 6 calendar months (partial current month included). Avg SV Hours is the mean of (epic actual-to-date minus epic planned-to-date) at each month end. Avg SV Date is a derived day-equivalent proxy: SV hours divided by each epic's own average daily planned rate (`planned_total_hours` / total planned days), documented as a proxy rather than a literal day count.
- The Portfolio Budget vs Actual Hours chart is an additional leadership-visibility chart requested for at-a-glance comparison; it uses positive `jira_original_estimate_hours`, or `planned_total_hours` when Jira original estimate is missing or zero, plus `total_actual_hours`. The shorter bar overlays the longer bar so an over-budget actual value does not hide the blue budget indicator.

## Business Rules

- Default scope is every canonical Jira epic in the active canonical run.
- Date filtering uses epic-level overlap only: `epic.start_date <= to_date` and `epic.due_date >= from_date`.
- Project filtering uses the canonical epic project only.
- Nested data is never date-filtered or project-trimmed after an epic is included.
- The top-level table uses compact rows, alternating row shading, visible row numbers, vertical scrolling, and sticky row-number/name columns so wide tables stay traceable while scrolling.
- Actual hours roll up from descendant `Sub-task` and `Bug Subtask` lifetime actuals. When individual `canonical_worklogs` rows are incomplete for a subtask because the refresh scope excludes older worklog dates, Epic Explorer still uses the canonical Jira issue `total_hours_logged` value for the hour total. Worklog-row detail continues to drive author, month, and dated drilldown analytics where those rows exist.
- Actual Complete Date is populated only for epics whose current status is `Resolved`, `Resolved!`, `Done`, `Closed`, `Complete`, or `Completed`. For those epics, use the later of the last descendant worklog date and `resolved_stable_since_date` when both exist, otherwise use whichever exists. Unresolved and reopened epics leave Actual Complete Date blank and use today as the SV reporting date, so their schedule variance continues to age even when no new worklogs are recorded.
- Top-level Planned vs Actual, SV Hours, and Est. Accuracy use the Jira epic's own planned dates and canonical epic `original_estimate_hours` as the primary planned basis. TK Budget, story estimates, and subtask estimates do not drive these calculations unless the epic original estimate is missing or zero.
- The SV Hours breakdown uses the same server-calculated values as the table KPI. Estimate fallback order is Jira Original Estimate, TK Budget, Story Estimates, then Subtask Estimates. The dialog never recalculates these values independently in the browser.
- Drawer month plan vs actual uses story/work-item original estimates only. It does not add subtask estimates to planned hours.
- Story estimates are distributed across each story/work-item `start_date`..`due_date` by daily capacity weight, then summed into calendar months. When `Assignee capacity after leaves` is selected, weekdays use the saved capacity profile's standard/Ramadan hours and holidays, and assigned stories subtract that assignee's leave hours for the day. If all capacity weights in a story range are zero, the story falls back to an even calendar-day split so the estimate is not dropped.
- Drawer month plan vs actual compares the story-derived monthly planned hours with descendant subtask worklogs by month and can be switched between bar and line chart views. The drawer also includes Story Plan Allocation so reviewers can trace each monthly planned value back to contributing stories and splits.
- Drawer Schedule Variance KPIs show date SV, hour SV, planned vs actual hours, planned vs actual delivery date, estimation accuracy, the last-three-month SV trend, and SV per assignee working on the epic.
- The drawer shows a planning alert when an estimated story/work item has no assignee. This warns reviewers that future capacity-based monthly allocation cannot use assignee capacity for those estimated items until Jira assignees are set.
- SV trend uses the latest three months available up to the epic status month. Trend is improving when the latest signed hour SV is higher than the first month in the window, declining when lower, and flat when unchanged.
- SV per assignee compares each story assignee's planned-to-date story estimate allocation with actual worklog hours on the epic. Stories without assignees are grouped as `Unassigned`.
- Drawer Gantt uses month, week, and day header rows. Each resource/day cell is green for actual work, orange for planned leave, red for unplanned leave, and mixed when more than one signal lands on the same resource/day.
- Leave overlays and the Leaves By Resource table are limited to people assigned subtasks in the selected epic; unrelated leave-only resources are not shown.
- Estimate quality counts subtasks whose original estimate equals their parent work-item estimate and subtasks whose logged actuals exceed their own original estimate. The section shows scorecards plus task-level detail rows.
- Task Completion displays readable bucket names and clickable counts. Clicking a count shows the contributing task details, including task name, work-item name, assignee, status, due date, and actual completion date.
- Work Item Effort uses work-item names in the chart label so reviewers do not need to remember Jira IDs.
- Resource Utilization keeps the table and also renders a bar-style chart summarizing epic hours and utilization percentage by resource.
- Team effort is mapped from Performance Settings team membership. Support effort is mapped from the Monthly Epic Plan support-team roster.

## Dependencies

| Source | Usage |
|---|---|
| `canonical_refresh_state` | Resolves the active successful canonical run. |
| `canonical_issues` | Supplies epic, work-item, subtask, bug-subtask hierarchy, status, assignee, dates, and original estimates. |
| `canonical_worklogs` | Supplies descendant worklog authors, dates, and hours for dated drilldowns and monthly/author analytics. |
| `canonical_issues.total_hours_logged` | Supplies lifetime Jira issue actual hours used as a fallback when it is greater than the scoped stored worklog rows for a descendant subtask or bug subtask. |
| `epics_management` | Supplies product, Jira URL, and TK budget metadata. |
| RLT canonical snapshot via `build_rlt_leave_snapshot()` | Supplies planned and unplanned leave cells for drawer Gantt and leave summary. |
| `performance_teams` | Maps worklog authors to team effort. |
| `support_team_config` | Identifies support resources who logged work on an epic. |

## Related Code

- `epic_explorer_service.py` - builds the API payload, rollups, nested hierarchy, drawer analytics, leave overlays, team effort, support effort, and resource utilization rows.
- `epic_explorer_report.html` - renders filters, table, nested expand/collapse drilldown, resizable analytics drawer, charts, Gantt cells, CSV export, and the client-side Executive Summary mini dashboard (epic picker, pinned-epic table, week-over-week quick-expand trend, month-over-month SV trend chart, and portfolio budget-vs-actual chart).
- `report_server.py` - registers `epic_explorer_report.html`, exposes `/api/epic-explorer/summary`, syncs the root HTML file into `report_html/`, and adds the report to page categorization.
- `tests/test_epic_explorer.py` - verifies rollups, filter semantics, route registration, page catalog registration, HTML sync, required UI controls, and the Executive Summary mini dashboard controls (epic picker, pinned table, quick-expand trend, month-over-month chart, leadership chart).

## Change Notes

- This report does not add or change database schema.
- `report_html/epic_explorer_report.html` is generated by the existing report HTML sync flow from the root `epic_explorer_report.html` source.
- Made each main-table SV Hours KPI clickable and added an accessible modal calculation breakdown backed by the API's exact schedule-variance inputs and results.
- Added the Executive Summary mini dashboard: an epic picker (dropdown + checkboxes + search), a pinned-epics table with quick-expand week-over-week schedule variance, month-over-month average schedule variance trend chart, and a portfolio budget-vs-actual leadership chart. This feature is entirely client-side (reuses the existing `/api/epic-explorer/summary` payload) and adds no new backend routes or schema changes. Pinned epic selection persists via browser `localStorage`.
- Updated the Executive Summary mini dashboard to show each pinned epic's status, keep search typing focused while filtering the picker list, and render the lower budget/actual bar above the larger bar so both values remain visible.
- Updated Executive Summary Budget values to use Jira Original Estimate rather than TK budget; the pinned table, SV Hours, and portfolio bar chart now compute from the same Jira estimate basis.
- Updated Executive Summary SV Hours to use visible Budget vs Actual delta instead of the detailed schedule-proration basis.
- Updated Executive Summary date formatting and alignment: weekly-trend chart labels now use `MMM-DD`, weekly-trend `Week Ending` uses `MMM-DD-YYYY`, pinned-table planned/actual date columns use `MMM-DD-YY`, and pinned-table text is centered (including Budget and Actual Hours).
- Adjusted Executive Summary layout so the Epic Name column is not centered, and increased Portfolio Budget vs Actual chart label width from 92px to 184px to reduce multi-line wrapping on long epic names.

## Business Logic

- The Executive Summary reuses the active epic summary payload already loaded for the page.
- The picker keeps a pending set for UI rendering, but every checkbox change immediately syncs that set into the pinned-epic dashboard state so charts and counts stay current.
- The picker search field is rendered once per menu render and filters the dedicated options container instead of rebuilding the whole menu on every keystroke.
- Apply is now a close-and-commit action, not the only moment when the dashboard changes.

## Business Cases

- Lets delivery leads keep a short list of important epics visible while they are triaging scope, schedule variance, and budget pressure.
- Removes the need to reapply a batch of selections just to see whether the dashboard changed.

## Examples

- Type multiple letters in the picker search field: the checkbox list filters in place and the typed search stays focused.
- Check one more epic in the picker: the pinned count increments right away and the portfolio chart adds one more bar pair.
- Uncheck a pinned epic: its row disappears from the mini dashboard immediately and the local storage snapshot updates.

## Explanations

- Open the epic picker, search for a key or name, and toggle the checkboxes you want pinned.
- Read the Status pill beside each pinned epic to distinguish resolved work from in-progress or blocked work.
- Expand an epic in the mini dashboard to see the line chart and data table for the same week-by-week proration values.
- The dashboard refreshes as soon as a checkbox changes, so you can see whether the new epic belongs in the leadership view before closing the picker.

## Front-end UI Fields

- Add epics: searchable checkbox dropdown; shows all loaded epics.
- Status: status pill in the pinned-epics table using the canonical epic status.
- Apply: closes the picker after the current live selection has already been saved.
- Clear all: removes every pinned epic and clears the dashboard immediately.
- Remove icon: unpins a single epic from the mini dashboard.

## Script Files

- `epic_explorer_report.html` - client-side Executive Summary picker, dashboard, and charts.
- `report_html/epic_explorer_report.html` - synced served copy for localhost and deployment.
- `tests/test_epic_explorer.py` - regression coverage for the report HTML and served route.

## Dependent & Impacted Files

- `report_server.py` - serves and syncs the report HTML.
- `tests/test_epic_explorer.py` - checks the picker and dashboard UI.
- `epic_explorer_service.py` - supplies the payload the picker and charts reuse.

## Table Schema

- `canonical_issues` - epic, story, subtask, and worklog hierarchy used by the report.
- `canonical_worklogs` - worklog hours and dates used by the schedule charts.
- `epics_management` - product metadata and TK budget values for the summary rows.

## Data Flow

1. `report_server.py` serves `epic_explorer_report.html`.
2. The page loads `/api/epic-explorer/summary`.
3. The Executive Summary picker renders from that payload only.
4. Checkbox changes update the pinned set, persist it to `localStorage`, and rerender the mini dashboard immediately.
5. The same payload feeds the leadership chart, the pinned table, and the week-over-week trend.
