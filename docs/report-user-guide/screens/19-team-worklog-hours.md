# Team Worklog Hours

## Business Cases
See each employee’s daily worklog coverage and RLT leave in a selected month, including dates without counted work.

## Business Logic
Each employee/date cell sums canonical Jira worklogs by author and date. A zero means no hours counted under the selected scope, not necessarily that the employee was expected to work. RLT leave taken is shown separately, using planned plus unplanned taken hours from the canonical daily leave snapshot. Leave worklogs are excluded by default; including them adds them to logged hours while still showing the separate leave indicator, so do not add the two values together. Dates without worklogs, including weekends and holidays, remain visible. The selected capacity profile identifies holiday/workday headings. Employee membership, resignations, support markings, and assigned-subtask scope follow Employee Capacity & Utilization.

## Examples
Alice logs 3 hours on an assigned subtask and 2 hours on a story on October 5, with 1 hour of RLT leave taken. All-work scope shows 5.00h and Leave 1.00h; assigned-subtask scope shows 3.00h and the same leave indicator. A date without matching logs shows 0.00h.

## Explanations
Select a month and team to compare people across days. Scroll horizontally for later dates. Green indicates positive counted hours and pale red indicates zero; a leave label appears where leave was taken. Total is the sum of logged hours across the month.

## Front-end UI Fields
Month defaults to the current month; capacity profile defaults to auto; Teams defaults to all except Process Team. Count logged hours from chooses all employee work or only work on their currently assigned subtasks. Include leaves in logged hours defaults off; Display resigned defaults off; Indicate support defaults on. Search employees narrows names. Show/Hide controls collapses filters. Download Excel exports the visible filtered matrix with separate numeric logged and RLT leave columns for every date. The utilization coloring control from the source report is hidden because utilization percentages are not displayed here.

## Script Files
`generate_team_worklog_hours_report.py` derives the shell from `generate_employee_capacity_utilization_report.py`; `team_worklog_hours.js` renders the matrix; `team_worklog_hours_report.html` is the generated shell. `report_server.py` registers the page, shared canonical data route and Excel endpoint. `shared-nav.js` exposes navigation. `run_html_only.py` rebuilds the shell.

## Dependent & Impacted Files
Shares the canonical payload, employee filters, profile selection and leave rules documented in [Employee Capacity & Utilization](16-employee-capacity-utilization.md). The report does not modify Jira or SQLite data.

## Table Schema
No new tables or schema changes. The shared canonical payload reads existing Jira exports and capacity configuration; see the [source report documentation](16-employee-capacity-utilization.md) for its data dependencies.

## Data Flow
The browser requests `/api/employee-capacity-utilization/data`; the server builds the existing canonical payload. The matrix groups filtered worklogs by author and date, combines matching daily leave totals, and renders all dates. Excel posts the displayed matrix to `/api/team-worklog-hours/export`, which creates a workbook and escapes spreadsheet formula-like text.

Weekend column headers and cells are shaded grey regardless of whether hours were logged; work and leave values remain visible. Open this report from the Team Worklog Hours button on Employee Capacity & Utilization.

When the selected month contains the viewer’s current local date, its header is labeled Today and the column is outlined in blue. Weekend cells keep their grey fill. Other months have no Today highlight.

Hover over any cell in an employee row for the availability hint, then click it (or focus it and press Enter or Space) to open the same availability drawer as Employee Capacity & Utilization. The drawer supports month navigation, a shareable link, live Jira worklog and RLT leave refresh, and the existing subtask and personal-worklog actions. Jira permissions and sign-in continue to govern those actions.

Selecting a date cell opens the availability drawer with **Log work** already active for that employee and date; the shareable link restores that selection. Selecting the employee name or monthly total opens the general availability drawer. Logging still requires the existing Jira sign-in and identity permissions, and no work is submitted automatically.

## Progressive live updates

The saved snapshot appears first. For each visible employee, the browser independently requests the current month’s Jira worklogs and RLT leave; it updates each result as it arrives. The row shows Work and Leave as updating, live with the check time, or update failed with either snapshot or last live values. Refresh live hours and leave retries the current month; an open visible report also refreshes every five minutes. A changed month cancels pending browser requests and discards their results. A rolling pool runs up to three employee requests at a time per browser, with three matrix-read slots per server process. All selected employees’ worklog requests finish (successfully or with a visible failure) before leave requests start. Free slots are filled immediately. Busy responses are retried up to twice using Retry-After (up to 30 seconds plus jitter); other failed reads remain visibly failed and do not replace values with zero. Live readings are in memory and do not update the canonical snapshot.

Live worklogs include all issue types, are grouped by Jira worklog author and started date, and fetch current Jira issue assignment for the assigned-subtask option. RLT leave uses the existing live leave redistribution logic by leave-subtask assignee. Jira account matching requires one unique employee and access to relevant issues and worklogs. Excel exports the displayed hours and a status column, including when data is stale. The matrix drawer retains its own independently refreshed availability calendar.

`team_worklog_live.py` and `/api/team-worklog-hours/live` provide the two independent read types. `worklog_tally.py` supplies paginated Jira reads and optionally includes assignee matching; `live_rlt_leave.py` calculates daily leave. No new database tables or migrations are required.

Changes to month, profile, teams, logged-hours scope, leave inclusion, resigned and support switches, and employee search immediately update the URL fragment. Copy or reload the URL to restore these report filters; availability drawer and selected-day links continue to use the same fragment. Filter edits replace the current browser history entry.

The progress line reports Hours checked, Leaves waiting/checked, and failures. Live responses recompute only the affected employee row and its Excel values; other rows, the table header, and table scroll position are preserved. Report-filter changes rebuild the displayed matrix as usual.
