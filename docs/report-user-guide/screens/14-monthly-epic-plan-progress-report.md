# Monthly Epic Plan Progress Report

## Business Logic

- The report compares TK planner epics for the selected month against Jira execution data and a canonical worklog snapshot.
- Epic scope starts from the selected month and the chosen epic mode, then adds unresolved brought-forward epics that still fall inside the overdue lookback threshold.
- The Estimate hierarchy stats only use the planned-this-month epic subset. Brought-forward overdue epics remain visible elsewhere in the report but do not inflate month-plan estimate bars.
- `Month Plan` is a reference bar based on the executive planned-hours total for the same visible month-planned epic set.
- `Epic Estimate`, `Story Estimate`, and `Subtask Estimate` each show a separate Jira original-estimate layer rather than replacing the main planned-hours calculation.
- `Subtask Logged` sums selected-month worklog hours on in-scope subtasks.
- `Story Overrun` on the main chart remains `max(sum(all selected-month subtask and bug-subtask logged hours for a story) - story original estimate, 0)`.
- Clicking an estimate bar opens the estimate detail drawer scoped to the currently visible report rows.
- The Story Overrun drawer stays at story level. Each row shows the story Jira original estimate, the story TK planned value when a planner-backed Jira link exists, and the story's summed subtask logged hours.
- The Story Overrun drawer excludes bug subtasks by default. The `Include bug subtasks` checkbox is drawer-only and does not change the main Story Overrun bar, percent, or count behind the drawer.
- When `Include bug subtasks` is off, the drawer recalculates each visible story row from regular subtasks only and hides rows whose non-bug overrun becomes zero.
- When `Include bug subtasks` is on, the drawer reverts to the combined regular-subtask plus bug-subtask totals already used by the main Story Overrun metric.
- The estimate detail drawer width can be resized from the left edge.
- Days are derived from `hours_per_day`, currently `8`.

## Business Cases

- Delivery leads use the report to compare month commitments against actual Jira execution for epics that should be progressing this month.
- Managers use the estimate hierarchy section to detect where planning detail disappears between epic, story, and subtask levels.
- The Story Overrun drawer is used during delivery reviews to see which stories exceeded their Jira original estimates and whether that overrun comes from normal implementation subtasks or bug-fix subtasks.
- The drawer-only bug toggle supports the business question “What is the overrun without defect rework?” without rewriting the headline Story Overrun stat used in the summary chart.

## Examples

- Example 1: A story has Jira original estimate `40 h`, regular subtasks logged `30 h`, and bug subtasks logged `18 h` in the selected month.
	- Main Story Overrun bar: `48 h - 40 h = 8 h`
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

The new bug-subtask toggle is intentionally local to the drawer. The report headline still answers “Which stories overran once all subtask work is counted this month?” The drawer toggle answers the narrower question “How much of that story overrun remains if bug subtasks are excluded?” This separation preserves stable top-line reporting while giving reviewers a diagnostic switch for defect-driven effort.

## Front-end UI Fields

| Area | Field | Type | Default | Behavior |
|---|---|---|---|---|
| Controls | Month | Month input | Current month | Defines the reporting month used for epic scope and selected-month worklogs. |
| Controls | Projects | Multi-select dropdown | All projects | Limits the payload to chosen project keys. |
| Controls | Employees | Hierarchical checkbox dropdown | Team-scope default selection | Limits workforce and leave calculations to chosen assignees. |
| Controls | Effort unit | Radio buttons | Hours | Switches visible numeric values between hours and days. |
| Controls | Epic scope | Segmented buttons | TK planner epics | Chooses planner-only, all epics, or all Jira epics mode. |
| Controls | Overdue lookback | Numeric input | 30 days | Sets how far back unresolved epics can be pulled forward into scope. |
| Controls | Date toggles | Toggle controls | Default mode | Applies start-date, due-date, or range client-side filtering. |
| Estimate hierarchy stats | Month Plan | Clickable bar | Calculated | Opens the drawer with visible in-month epic rows and their planned-vs-logged context. |
| Estimate hierarchy stats | Epic Estimate | Clickable bar | Calculated | Opens epic Jira original-estimate rows. |
| Estimate hierarchy stats | Story Estimate | Clickable bar | Calculated | Opens story Jira original-estimate rows. |
| Estimate hierarchy stats | Subtask Estimate | Clickable bar | Calculated | Opens subtask Jira original-estimate rows. |
| Estimate hierarchy stats | Subtask Logged | Clickable bar | Calculated | Opens subtask selected-month logged-hour rows. |
| Estimate hierarchy stats | Story Overrun | Clickable bar | Calculated | Opens story-level overrun rows. Main bar still includes both regular subtasks and bug subtasks. |
| Estimate detail drawer | Jira | Link/button column | Derived | Opens the work item in Jira when a URL is available. |
| Estimate detail drawer | Work item | Text column | Derived | Shows summary and work-item type. |
| Estimate detail drawer | Parents | Link chips | Derived | Shows the parent story and epic links where relevant. |
| Estimate detail drawer | Planned dates | Text column | Derived | Displays start and due dates as `start → due`. |
| Estimate detail drawer | Original estimate | Numeric column | Derived | Shows Jira original estimate in the selected unit. |
| Estimate detail drawer | TK planned | Numeric column | Blank when missing | Shows planner-backed TK planned effort for matching epic/story Jira links. |
| Estimate detail drawer | Logged hours | Numeric column | Derived | For Story Overrun rows, uses non-bug totals by default and combined totals when the bug toggle is enabled. |
| Estimate detail drawer | Overrun | Numeric column | Derived | For Story Overrun rows, recalculated from the currently displayed logged totals. |
| Estimate detail drawer | Include bug subtasks | Checkbox | Unchecked | Available only in the Story Overrun drawer. Recalculates drawer rows and summary chips without changing the main chart metric. |
| Estimate detail drawer | Summary chips | Read-only chips | Calculated | Shows work-item count, original estimate, TK planned when present, logged, and overrun for the currently displayed drawer rows. |

## Script Files

- `monthly_epic_plan_progress_service.py` — builds the monthly payload, estimate rollups, and Story Overrun detail rows including split regular-subtask and bug-subtask logged totals.
- `monthly_epic_plan_progress_report.html` — renders filters, estimate bars, the detail drawer, the drawer-only bug toggle, table view, and Gantt view.
- `report_server.py` — serves `/api/monthly-epic-plan-progress/summary` and syncs the canonical report HTML into `report_html/` for localhost serving.
- `tests/test_monthly_epic_plan_progress.py` — covers payload generation, HTML presence checks, and Story Overrun detail-row regression cases.

## Dependent & Impacted Files

- `report_server.py` depends on this module because it exposes the summary API consumed by the page.
- `tests/test_monthly_epic_plan_progress.py` is impacted whenever estimate-rollup payload fields or drawer markup change.
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
5. The service reads matching epic, story, subtask, and bug-subtask rows from `canonical_issues` and selected-month worklogs from `canonical_worklogs`.
6. The service builds top-line totals, estimate hierarchy rollups, and detail rows by metric.
7. For Story Overrun rows, the service keeps the headline combined overrun fields and also stores regular-subtask-only logged and overrun values for drawer filtering.
8. The browser renders the bar chart from the unchanged main rollup fields.
9. When the user opens the Story Overrun drawer, the browser starts from the story-level detail rows, defaults the bug toggle to unchecked, recalculates displayed logged/overrun values from the regular-subtask-only fields, and removes rows that no longer overrun.
10. If the user enables `Include bug subtasks`, the browser re-renders the same rows using the combined totals already supplied by the backend.
