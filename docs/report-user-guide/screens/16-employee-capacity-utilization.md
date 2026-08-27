# Employee Capacity & Utilization

## Purpose

This standalone report presents employee capacity, profile holidays, personal leave, booked subtask estimates, logged hours, and utilization without Employee Performance scoring or dashboard content.

## Data and calculations

- It loads live from `/api/employee-capacity-utilization/data` whenever the page opens. The endpoint reads the server's canonical database, using the active successful `canonical_issues` snapshot and newest persisted `canonical_worklogs` run. The deployed HTML does not rely on build-time database contents.
- Capacity comes from the selected capacity profile. Weekday official leaves in that profile are excluded from capacity before Availability is calculated; Availability is then capacity less planned and unplanned leave. Weekend official-leave dates remain visible in the Official Leaves count but do not reduce hours because they are already non-working days.
- Booked Manhours is the original estimate of assigned subtasks overlapping the selected month. Logged Hours has two clearly labelled scopes: **All work logged by employee** counts every in-month Jira worklog authored by that employee; **Work logged on their assigned subtasks** counts only time on subtasks currently assigned to the same employee. The narrower scope excludes stories, tasks, bugs, and work items assigned to somebody else. RLT-board leave worklogs are excluded by default because leave hours are already reported in Leaves Taken; the user can explicitly include them when needed.
- The final Grand Total row sums the displayed employees. Official Leaves shows the selected profile's holiday-day count per employee and the corresponding resource-day total in the Grand Total row.

## Filters

The filter bar is divided into two purposeful rows. **Refine the report** contains the primary scope choices—Month, Capacity Profile, Teams, and the prominently highlighted Logged Hours source. **Display & inclusion** contains presentation and visibility preferences—color formatting, leave-worklog inclusion, resigned employees, and Support badges. All changes apply immediately; the status pill confirms this behaviour.

| Field | Default | Effect |
| --- | --- | --- |
| Month | Current month | Limits worklogs, leave, and booked-subtask date overlap. |
| Capacity profile | Auto | Selects the capacity calendar. |
| Teams | All except Process Team | Includes unassigned employees and employees in selected configured teams; Process-Team-only resources are hidden by default. Changes apply immediately. |
| Count logged hours from | All work logged by employee | **All work logged by employee** counts any selected-month worklog by the employee, regardless of work-item type or assignee. **Work logged on their assigned subtasks** counts only subtask worklogs where the worklog author and current assignee are the same person. The table header, status line, detail drawer, and Excel export retain the selected scope. |
| Include leaves in logged hours | Off | Includes RLT-project and `RLT-*` worklogs in Logged Hours, utilization, color formatting, drawers, Grand Total, and Excel. Off prevents leave hours from being counted twice. |
| Color formatting | Utilization only; A=50%, B=80%, C=100% | Colors utilization from 0–A red, above A–B orange, and above B–C green. Values above C remain unfilled. The user can apply the color to the Utilization cell only or the entire employee row. Rules apply immediately and persist in browser storage. |
| Display resigned | Off | Includes resigned resources. |
| Indicate support | On | Adds a Support chip for `support_team_config` members. |

There is no Apply or Refresh button. Every filter change redraws the table immediately. If canonical data cannot be loaded, the report shows an explicit error rather than a misleading table of zero hours.

## Detail drawer

Every employee and metric value, including Grand Total values, is clickable and keyboard accessible. Activating a value opens a right-side drawer containing the source records for the active month and filters. Clicking Employee Name opens that employee's hour logs. The drawer title area also shows the active Logged Hours scope. In Logged Hours, the Type column uses a blue Google person icon for Stories, a purple bolt for Epics, a green plus for Subtasks, and a red bug for Bug Subtasks. A red Google warning icon marks every other work-item type, explaining that it is excluded by the assigned-subtask scope. The drawer also lists the Jira work-item link, title, Epic link, Epic name, worklog author, assigned resource, worklog date, and hours. Booked Manhours lists the original-estimate source subtasks. Leave, calendar capacity, availability, and utilization cells each use a metric-specific table.

The drawer can be widened or narrowed by dragging its left resize handle. The width persists in browser storage. It closes through its close button, the backdrop, or Escape, and returns focus to the triggering table cell.

## Excel export

Download Excel exports the complete dataset for the active Month, Capacity Profile, Teams, Logged Hours scope, leave-worklog inclusion, and resigned-resource setting. The workbook contains Export Info, Summary, Worklogs, Booked Subtasks, Leave Records, Capacity Calendar, and Employees worksheets. Work item and Epic URLs are active Excel hyperlinks. Each data sheet uses frozen headers, filters, table styling, readable column widths, and numeric hour/utilization columns.

## Related code

- `generate_employee_capacity_utilization_report.py`
- `employee_capacity_utilization_export.py`
- `report_server.py`
- `run_html_only.py`
- `shared-nav.js`

## Change notes

- The report now uses production runtime data instead of the canonical payload embedded when the HTML was generated.
- Filters were reorganized into a responsive control panel with switch controls, a team-count summary, and a sticky, readable data table.
- The Teams dropdown uses compact aligned checkboxes, single-line team labels, sticky selection actions, a bounded scroll area, and closes on outside click or Escape.
- Utilization thresholds and row-versus-column color scope are user-configurable, validated as `0 ≤ A < B < C ≤ 500`, and saved per browser.
- Every summarized value opens a resizable, metric-specific raw-data drawer.
- Employee Name opens hour logs, worklogs include Epic context, and the current filtered dataset can be downloaded as a structured multi-sheet Excel workbook.
- RLT-board leave worklogs are excluded from Logged Hours by default and can be restored instantly with Include leaves in logged hours; the same selection is applied to the table, drawer, totals, utilization, colors, and Excel export.
- Official-leave calendar calculations use local noon for date matching, preventing UTC date shifts from causing weekday official leave to remain in availability for users east of UTC.
