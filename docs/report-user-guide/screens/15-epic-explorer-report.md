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

## Fields And Validations

| Area | Field | Type | Default | Behavior |
|---|---|---|---|---|
| Filters | From | Date input | Blank | When both dates are set, includes only epics whose Jira epic `start_date` to `due_date` overlaps the range. |
| Filters | To | Date input | Blank | Pairs with From for overlap filtering. If only one side reaches the API, the server treats it as a single-day filter. |
| Actions | Apply dates / Clear dates | Buttons | - | Apply or clear only the date overlap filter. Project selections remain as-is. |
| Filters | Projects | Checkbox menu | All projects | Shows project display names when Epics Planner metadata is available. Includes only epics whose canonical epic `project_key` matches selected projects after Apply projects is clicked. |
| Actions | Apply projects / Clear projects | Buttons | - | Apply or clear only the project filter. Date filters remain as-is. |
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
| Table | Total Actual Hours | Numeric hours | 0 | Lifetime sum of worklog hours on descendant subtasks and bug subtasks. |
| Table | Actual Complete Date | Date | Blank when missing | Later of descendant subtask last worklog date and epic `resolved_stable_since_date`. |
| Table | Planned vs Actual Hours | Numeric pair | 0 / 0 | Compares planned total hours with lifetime actual hours. Planned total uses TK budget first, then Jira epic original estimate, then story estimate total, then subtask estimate total. |
| Table | Planned vs Actual Delivery | Date pair | Blank | Compares planned epic due date with actual complete date. |
| Table | SV Date | Signed days | Blank when no due date | Schedule variance by date: planned due date minus actual complete date, or current date for incomplete epics. Negative values mean behind schedule. |
| Table | SV Hours | Signed hours and percent | Blank when no planned-to-date hours | Schedule variance by effort: actual-to-date hours minus planned-to-date hours. Negative values mean behind schedule. |
| Table | Est. Accuracy | Percent | Blank when no actual hours | Estimated hours divided by actual hours multiplied by 100. The ideal range is 85% to 115%; below 70% indicates a broken estimation process. |
| Table | Epic Status | Status pill | Derived | Canonical epic `status`. |
| Table | Headcount | Integer | 0 | Distinct worklog authors on descendant subtasks and bug subtasks. |

## Business Rules

- Default scope is every canonical Jira epic in the active canonical run.
- Date filtering uses epic-level overlap only: `epic.start_date <= to_date` and `epic.due_date >= from_date`.
- Project filtering uses the canonical epic project only.
- Nested data is never date-filtered or project-trimmed after an epic is included.
- The top-level table uses compact rows, alternating row shading, visible row numbers, vertical scrolling, and sticky row-number/name columns so wide tables stay traceable while scrolling.
- Actual hours roll up from descendant `Sub-task` and `Bug Subtask` worklogs only.
- Actual complete date mirrors existing completion logic: use the later of last logged date and resolved-stable-since when both exist, otherwise use whichever exists.
- Drawer month plan vs actual compares month-bucketed planned original estimates with worklogs by month.
- Drawer month plan vs actual compares month-bucketed planned original estimates with worklogs by month and can be switched between bar and line chart views.
- Drawer Schedule Variance KPIs show date SV, hour SV, planned vs actual hours, planned vs actual delivery date, estimation accuracy, the last-three-month SV trend, and SV per assignee working on the epic.
- SV trend uses the latest three months available up to the epic status month. Trend is improving when the latest signed hour SV is higher than the first month in the window, declining when lower, and flat when unchanged.
- SV per assignee compares each assignee's planned-to-date estimate allocation with actual worklog hours on the epic.
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
| `canonical_worklogs` | Supplies descendant worklog authors, dates, and hours. |
| `epics_management` | Supplies product, Jira URL, and TK budget metadata. |
| RLT canonical snapshot via `build_rlt_leave_snapshot()` | Supplies planned and unplanned leave cells for drawer Gantt and leave summary. |
| `performance_teams` | Maps worklog authors to team effort. |
| `support_team_config` | Identifies support resources who logged work on an epic. |

## Related Code

- `epic_explorer_service.py` - builds the API payload, rollups, nested hierarchy, drawer analytics, leave overlays, team effort, support effort, and resource utilization rows.
- `epic_explorer_report.html` - renders filters, table, nested expand/collapse drilldown, resizable analytics drawer, charts, Gantt cells, and CSV export.
- `report_server.py` - registers `epic_explorer_report.html`, exposes `/api/epic-explorer/summary`, syncs the root HTML file into `report_html/`, and adds the report to page categorization.
- `tests/test_epic_explorer.py` - verifies rollups, filter semantics, route registration, page catalog registration, HTML sync, and required UI controls.

## Change Notes

- This report does not add or change database schema.
- `report_html/epic_explorer_report.html` is generated by the existing report HTML sync flow from the root `epic_explorer_report.html` source.
