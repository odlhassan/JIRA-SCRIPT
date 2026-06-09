# Monthly Epic Plan Progress Report

## Business Logic

- The report compares TK planner epics for the selected month against Jira execution data and a canonical worklog snapshot.
- Epic scope starts from the selected month and the chosen epic mode, then adds unresolved brought-forward epics that still fall inside the overdue lookback threshold.
- The Estimate hierarchy stats only use the planned-this-month epic subset. Brought-forward overdue epics remain visible elsewhere in the report but do not inflate month-plan estimate bars.
- `Month Plan` is a reference bar based on the executive planned-hours total for the same visible month-planned epic set.
- `Epic Estimate`, `Story Estimate`, and `Subtask Estimate` each show a separate Jira original-estimate layer rather than replacing the main planned-hours calculation.
- `Subtask Logged` sums selected-month worklog hours on in-scope subtasks. The top-level `Include Bug Subtasks` control is on by default; when it is turned off, Bug Subtask rows are removed before planned-hour, logged-hour, child-row, Gantt/worklog-marker, estimate-rollup, and project-card totals are calculated.
- `Story Overrun` on the main chart is `max(sum(selected-month included-subtask logged hours for a story) - story original estimate, 0)`. With the default bug toggle on, included subtasks means regular subtasks plus Bug Subtask issues. With the toggle off, included subtasks means regular subtasks only.
- Clicking an estimate bar opens the estimate detail drawer scoped to the currently visible report rows.
- The Story Overrun drawer stays at story level. Each row shows the story Jira original estimate, the story TK planned value when a planner-backed Jira link exists, and the story's summed subtask logged hours.
- The Story Overrun drawer still has a drawer-local `Include bug subtasks` diagnostic switch for overrun story details. The page-level `Include Bug Subtasks` switch controls the server payload first; the drawer switch can only include bug-subtask rows that are already present in that payload.
- The Epics table shows `TK Epic Budget`, `Month Plan`, `Month Actual`, and `Total Actual` as separate effort columns. `TK Epic Budget` comes from the epic-level TK-budgeted man-days when available, falling back to the epic plan man-days.
- Epics table project chips and row stripes use the managed project color configured in Projects settings.
- The estimate detail drawer width can be resized from the left edge.
- Days are derived from `hours_per_day`, currently `8`.

### Workforce: Team Roster, Employee Stats, and Resource Planning

- The **Team Roster** drawer is the single source of truth for which employees the workforce numbers describe. By default it leaves only the **active development team** selected; support-team members, employees whose resignation date is before the selected month, and process-team resources are unchecked.
- **Include Support Team** toggle:
	- **On** → the user wants combined active dev + active support stats; support members become eligible for selection.
	- **Off** (default) → only the active dev team is in scope; support members are force-unchecked.
- **Employee Stats** cards (head count, capacity, leaves, availability) reflect the roster selection: head count shows the selected count, and capacity/leaves are recomputed server-side as `team_capacity × (selected ÷ profile headcount)` and the sum of selected members' planned leave. Manual selections are preserved when the user reopens either the header employee dropdown or Team Roster and clicks Apply; resignation records do not block explicit selection.
- **Resource Planning** (Total / Dev / Support resources) now mirrors the same roster selection instead of a fixed `all − process − resigned` formula. The panel is computed client-side from the selected members so it always stays consistent with Employee Stats:
	- Per-member capacity is uniform: `team_capacity_hours ÷ employee_count_profile` (the same `per_person_capacity_hours` basis the server uses).
	- **Total** = selected members × per-person capacity; planned leave summed over selected members; availability = capacity − leave.
	- **Support** = the subset of selected members that belong to the support team (same per-person basis).
	- **Dev** = Total − Support for every metric, so Dev head count never includes a member the user unchecked.
	- When Include Support Team is off, no support member is selected, so the Support Resources group is hidden and **Total = Dev**.
- On first load with no explicit selection, Resource Planning still mirrors the roster default (active dev only, support shown only when Include Support Team is applied on), rather than the full organisation breakdown. A resignation dated before the selected month is treated as inactive for that month; a resignation dated inside or after the selected month remains selectable and active for that month.
- The detailed **Technical Support Team** table below Employee Stats continues to list the full support roster as a reference; only the Resource Planning *Support Resources* card is selection-aware.

## Business Cases

- Delivery leads use the report to compare month commitments against actual Jira execution for epics that should be progressing this month.
- Managers use the estimate hierarchy section to detect where planning detail disappears between epic, story, and subtask levels.
- Delivery leads use the top-level `Include Bug Subtasks` toggle to reconcile EPR monthly totals with Jira exports that include defect subtasks in the downloaded issue set.
- The Story Overrun drawer is used during delivery reviews to see which stories exceeded their Jira original estimates and whether that overrun comes from normal implementation subtasks or bug-fix subtasks.
- The drawer-local bug toggle supports the narrower business question “What is the overrun without defect rework?” within the already-loaded report scope.

## Examples

- Example 1: A story has Jira original estimate `40 h`, regular subtasks logged `30 h`, and bug subtasks logged `18 h` in the selected month.
	- Top-level `Include Bug Subtasks` on: executive actual includes `48 h`, and the main Story Overrun bar is `48 h - 40 h = 8 h`
	- Top-level `Include Bug Subtasks` off: executive actual includes `30 h`, and the main Story Overrun bar is `max(30 h - 40 h, 0) = 0 h`
	- Story Overrun drawer, toggle off: logged `30 h`, overrun `0 h`, row hidden because non-bug work did not exceed the story estimate
	- Story Overrun drawer, toggle on: logged `48 h`, overrun `8 h`, row shown
- Example 2: A story has Jira original estimate `16 h`, TK planned `24 h`, and regular subtasks logged `20 h` with no bug subtasks.
	- Drawer row always shows original estimate `16 h`, TK planned `24 h`, logged `20 h`, overrun `4 h`
- Example 3: An epic is brought forward from a prior month.
	- It can still appear in the executive summary and epic table
	- It does not contribute to Estimate hierarchy stats such as Month Plan, Story Estimate, or Story Overrun

## Explanations

The screen starts with month, project, employee, effort-unit, epic-scope, overdue-lookback, and date-toggle filters. The server uses those filters to assemble the current epic population from planner rows and canonical Jira data. It then calculates summary totals, estimate hierarchy rollups, project cards, workforce information, and table/Gantt rows.

The estimate hierarchy bar chart is a fast comparison layer. Each bar opens a drawer with the work items that contributed to the clicked metric. For Story Overrun, the drawer lists stories rather than individual subtasks so reviewers can compare a story's Jira estimate, planner-backed TK plan, and the selected-month logged hours in one row.

The page-level bug-subtask toggle is a server-side scope switch. It defaults on because Jira recon exports commonly include both `Sub-task` and `Bug Subtask` issue types. Turning it off re-fetches the payload with Bug Subtask issue keys excluded from the subtask-to-epic map, so planned hours, actual worklogs, estimate hierarchy stats, child rows, and Gantt worklog indicators all stay internally consistent.

Jira CSV reconciliation has one important caveat: the downloaded `Σ Time Spent` column is Jira's aggregate lifetime time-spent value for each returned issue, while the EPR `Logged this month` KPI uses `canonical_worklogs.started_date` inside the selected month. A Jira CSV filtered by issue start/due dates can therefore be higher than EPR if the returned issues contain worklogs from other months, or if the report is in `TK EPICS` scope while the CSV includes all Jira issues. To compare like-for-like, use `ALL JIRA EPICS`, keep `Include Bug Subtasks` on, and compare against a Jira export built from worklog dates for the same calendar month.

## Front-end UI Fields

| Area | Field | Type | Default | Behavior |
|---|---|---|---|---|
| Controls | Month | Month input | Current month | Defines the reporting month used for epic scope and selected-month worklogs. |
| Controls | Projects | Multi-select dropdown | All projects | Limits the payload to chosen project keys. |
| Controls | Employees | Hierarchical checkbox dropdown | Team-scope default selection | Limits workforce and leave calculations to chosen assignees. |
| Controls | Effort unit | Radio buttons | Hours | Switches visible numeric values between hours and days. |
| Controls | Epic scope | Segmented buttons | TK planner epics | Chooses planner-only, all epics, or all Jira epics mode. |
| Controls | Overdue lookback | Numeric input | 30 days | Sets how far back unresolved epics can be pulled forward into scope. |
| Controls | Include Bug Subtasks | Checkbox toggle | Checked | Re-fetches the report with Bug Subtask planned and logged hours included or excluded. Checked by default to match Jira recon exports that include bug subtasks. |
| Controls | Date toggles | Toggle controls | Default mode | Applies start-date, due-date, or range client-side filtering. |
| Estimate hierarchy stats | Month Plan | Clickable bar | Calculated | Opens the drawer with visible in-month epic rows and their planned-vs-logged context. |
| Estimate hierarchy stats | Epic Estimate | Clickable bar | Calculated | Opens epic Jira original-estimate rows. |
| Estimate hierarchy stats | Story Estimate | Clickable bar | Calculated | Opens story Jira original-estimate rows. |
| Estimate hierarchy stats | Subtask Estimate | Clickable bar | Calculated | Opens subtask Jira original-estimate rows. |
| Estimate hierarchy stats | Subtask Logged | Clickable bar | Calculated | Opens subtask selected-month logged-hour rows. |
| Estimate hierarchy stats | Story Overrun | Clickable bar | Calculated | Opens story-level overrun rows. Main bar follows the top-level Bug Subtask inclusion setting. |
| Estimate detail drawer | Jira | Link/button column | Derived | Opens the work item in Jira when a URL is available. |
| Estimate detail drawer | Work item | Text column | Derived | Shows summary and work-item type. |
| Estimate detail drawer | Parents | Link chips | Derived | Shows the parent story and epic links where relevant. |
| Estimate detail drawer | Planned dates | Text column | Derived | Displays start and due dates as `start → due`. |
| Estimate detail drawer | Original estimate | Numeric column | Derived | Shows Jira original estimate in the selected unit. |
| Estimate detail drawer | TK planned | Numeric column | Blank when missing | Shows planner-backed TK planned effort for matching epic/story Jira links. |
| Estimate detail drawer | Logged hours | Numeric column | Derived | For Story Overrun rows, uses non-bug totals by default and combined totals when the drawer bug toggle is enabled. |
| Estimate detail drawer | Overrun | Numeric column | Derived | For Story Overrun rows, recalculated from the currently displayed logged totals. |
| Estimate detail drawer | Include bug subtasks | Checkbox | Unchecked | Available only in the Story Overrun drawer. Recalculates drawer rows and summary chips inside the current page-level bug-subtask scope. |
| Estimate detail drawer | Summary chips | Read-only chips | Calculated | Shows work-item count, original estimate, TK planned when present, logged, and overrun for the currently displayed drawer rows. |
| Estimate hierarchy stats | Logged this month (table_chart icon) | Icon button | — | Opens the Worklog Detail drawer showing per-subtask planned and logged hours for a selectable date range. |
| Worklog Detail drawer | From date / To date | Date pickers | Month start / month end | Select a precise start and end date for the worklog query. Click Apply to re-fetch data for the new range. |
| Worklog Detail drawer | Apply button | Button | — | Triggers a new API request to `/api/monthly-epic-plan-progress/worklog-detail` with the selected From/To dates. |
| Worklog Detail drawer | Download CSV | Button | — | Exports the full drawer data to a CSV file. Each per-worklog entry becomes its own row; subtask metadata is on the first worklog row only. Filename includes the active date range. |
| Worklog Detail drawer | Search | Search input | Empty | Filters displayed rows client-side by Jira key, summary, story key, or epic key. |
| Worklog Detail drawer | Summary chips | Read-only chips | Calculated | Shows subtask count, estimated hours, logged-in-range hours, and bug-subtask count. |
| Worklog Detail drawer | Header totals | Read-only cards | Calculated | Two total cards: "Logged in range" (sum of `month_logged_hours` for the queried date span) and "Total ever logged" (sum of `total_hours_logged` lifetime per subtask). |
| Worklog Detail drawer | Jira Key | Link column | Derived | Clickable link opening the Jira issue in a new tab when a URL is available. |
| Worklog Detail drawer | Type | Chip column | Derived | Shows `Sub-task` or `Bug Subtask`. |
| Worklog Detail drawer | Summary | Text column | Derived | Jira issue summary. |
| Worklog Detail drawer | Story / Epic | Text columns | Derived | Parent story and epic keys. |
| Worklog Detail drawer | Start Date | Date column | Derived | Jira `start_date` from `canonical_issues`. |
| Worklog Detail drawer | Due Date | Date column | Derived | Jira `due_date` from `canonical_issues`. |
| Worklog Detail drawer | Estimated | Numeric column | Derived | Jira original estimate in the selected unit. |
| Worklog Detail drawer | Logged | Numeric column / expand button | Derived | Worklog hours within the selected From/To date range. When individual worklog entries exist, the cell is a clickable button; clicking expands an inline nested table showing each worklog's date, author, and hours. |

| Epics table | TK Epic Budget | Numeric column | Derived | Shows epic-level TK budget in the selected unit. Blank for rows without planner-backed TK budget. |
| Epics table | Month Plan / Month Actual / Total Actual | Numeric columns | Derived | Show selected-month plan, selected-month worklogs, and total logged work across the epic. |
| Team Roster drawer | Include Support Team | Checkbox | Unchecked | When on, support members are selectable and feed Employee Stats + Resource Planning; when off, support members are force-unchecked (dev-only scope). |
| Team Roster drawer | Member / team checkboxes | Checkboxes | Active dev selected, support and before-month resignations unchecked | Selecting/clearing members defines the workforce scope. `Apply selection` re-fetches the payload with the chosen assignees and preserves manually selected employees, including employees with resignation records. |
| Employee Stats | Head Count / Capacity / Leaves / Availability | Read-only cards | Calculated | Reflect the roster selection (selected count, scaled capacity, selected members' planned leave). |
| Resource Planning | Total / Dev / Support resources | Read-only cards | Calculated | Mirror the roster selection. Total and Support are computed from selected members at a uniform per-person capacity; Dev = Total − Support. Support group hidden when no support member is selected. |

## Script Files

- `monthly_epic_plan_progress_service.py` — builds the monthly payload, estimate rollups, top-level Bug Subtask inclusion/exclusion, Story Overrun detail rows, month-aware resignation eligibility for workforce roster rows, and the standalone `build_worklog_detail_for_range()` function used by the dedicated worklog-detail endpoint.
- `monthly_epic_plan_progress_report.html` — renders filters, the top-level `Include Bug Subtasks` toggle, estimate bars, detail drawers, the Worklog Detail drawer with From/To date pickers and Apply button, table view, Gantt view, the Team Roster drawer, Employee Stats cards, month-aware resigned/inactive labels, and the selection-aware Resource Planning panel. The main `<style>` block must close before `<body>` so browser parsing does not treat the report body as CSS text.
- `report_server.py` — serves `/api/monthly-epic-plan-progress/summary` (includes `include_bug_subtasks` param) and the new `/api/monthly-epic-plan-progress/worklog-detail` endpoint (`from_date`, `to_date`, `include_bug_subtasks` params). Syncs canonical report HTML into `report_html/`.
- `tests/test_monthly_epic_plan_progress.py` — covers payload generation, HTML presence, Story Overrun regression, bug-subtask toggle, worklog detail payload, `build_worklog_detail_for_range` with custom dates, month-aware roster resignation eligibility, manual employee-selection persistence, and the HTML structural guard that `</style>` appears before `<body>`.

## Dependent & Impacted Files

- `report_server.py` depends on this module because it exposes the summary API consumed by the page.
- `tests/test_monthly_epic_plan_progress.py` is impacted whenever estimate-rollup payload fields, bug-subtask inclusion rules, workforce roster selection rules, resignation eligibility, or drawer/top-level toggle markup change.
- `report_html/monthly_epic_plan_progress_report.html` is a served copy produced by the sync flow and reflects changes from the canonical root HTML file.
- Planner-backed Jira link mappings from `epics_management` affect whether the drawer can display story or epic `TK planned` values.
- `support_center_service.py` reuses this module's support-team capacity helpers (`build_workforce_month_payload`, `HOURS_PER_DAY`, `_month_bounds`) read-only to compute "hours available" for the Support Center report. See `SUPPORT_CENTER_REPORT.md`.

## Table Schema

| Table | Columns Used | Notes |
|---|---|---|
| `canonical_refresh_state` | `active_run_id`, `last_success_run_id`, `updated_at_utc` | Identifies the Jira snapshot that should drive the current payload. |
| `canonical_issues` | `run_id`, `issue_key`, `project_key`, `issue_type`, `summary`, `status`, `assignee`, `start_date`, `due_date`, `resolved_stable_since_date`, `original_estimate_hours`, `total_hours_logged`, `parent_issue_key`, `story_key`, `epic_key` | Supplies epic, story, subtask, and bug-subtask hierarchy plus Jira original estimates and dates. |
| `canonical_worklogs` | `run_id`, `worklog_id`, `issue_key`, `project_key`, `worklog_author`, `issue_assignee`, `started_date`, `hours_logged` | Supplies selected-month logged hours for in-scope subtasks. |
| `epics_management` | `id`, `epic_key`, `project_key`, `project_name`, `product_category`, `component`, `epic_name`, `delivery_status`, `jira_url`, `epic_plan_json`, planner phase JSON columns | Supplies TK planned effort and Jira-link mappings for epic/story rows shown in the drawer. |

## Data Flow

1. The browser loads `monthly_epic_plan_progress_report.html` and gathers the current month, project, employee, unit, epic-mode, overdue-lookback, and date-toggle settings.
2. The page calls `/api/monthly-epic-plan-progress/summary` on `report_server.py`.
3. `report_server.py` delegates to `monthly_epic_plan_progress_service.py`.
4. The service reads the active canonical snapshot from `canonical_refresh_state`, then loads planner-backed epics from `epics_management`.
5. The service reads matching epic, story, subtask, and bug-subtask rows from `canonical_issues` and selected-month worklogs from `canonical_worklogs`; if `include_bug_subtasks=0`, Bug Subtask issue keys are excluded before the worklog lookup and rollup calculations.
6. The service builds top-line totals, estimate hierarchy rollups, and detail rows by metric.
7. For Story Overrun rows, the service keeps the current-scope combined overrun fields and also stores regular-subtask-only logged and overrun values for drawer filtering.
8. The browser renders the bar chart from the current-scope main rollup fields.
9. Changing the page-level `Include Bug Subtasks` toggle calls the summary API again with `include_bug_subtasks=1` or `0`.
10. When the user clicks the `table_chart` icon on the Logged This Month card, the Worklog Detail drawer opens with From/To date pickers pre-filled to the loaded month's first and last day, then immediately calls `/api/monthly-epic-plan-progress/worklog-detail?from_date=…&to_date=…&include_bug_subtasks=…`. The drawer renders per-subtask rows with 9 columns: Jira Key, Type, Summary, Story, Epic, Start Date, Due Date, Estimated, and Logged. Two header total cards show the sum of Logged in range and Total ever logged. Each Logged cell is a clickable expand button when individual worklog entries exist; clicking toggles an inline nested table of worklog date, author, and hours.
11. Changing the From or To date and clicking Apply re-calls the worklog-detail endpoint; the search box filters already-fetched rows client-side without a new network request.
12. Clicking "Download CSV" exports all drawer rows to a file named `worklog_detail_<from>_to_<to>.csv`. Each per-worklog entry becomes a separate row; subtask metadata repeats only on the first row for that subtask.
12. When the user opens the Story Overrun drawer, the browser starts from the story-level detail rows, defaults the drawer bug toggle to unchecked, recalculates displayed logged/overrun values from the regular-subtask-only fields, and removes rows that no longer overrun.
13. If the user enables the drawer-local `Include bug subtasks`, the browser re-renders the same rows using the combined totals already supplied by the backend for the current page-level scope.
14. For workforce numbers, the service returns per-member data in `employee_tree` (name, static `resigned`, month-aware `active_in_month` / `resigned_for_month`, planned `leave_hours`), the support roster in `support_team.member_rows`, `team_capacity_hours`, `employee_count_profile`, and the active selection (`selected_assignees`, `assignee_filter_active`).
15. `renderWorkforce` and `renderSupportTeam` call `computeResourcePlanningState`, which derives the effective selected set (from `selected_assignees`, or the active dev-only default when no filter is active), computes Total and Support buckets at a uniform per-person capacity, and `renderResourceSummary` then renders Dev = Total − Support. Changing the header employee dropdown or Team Roster selection and clicking Apply re-fetches the payload and re-renders all workforce panels consistently without reverting newly checked members.
