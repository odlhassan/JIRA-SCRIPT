# Employee Capacity & Utilization

## Purpose

This standalone report presents employee capacity, profile holidays, personal leave, booked subtask estimates, logged hours, and utilization without Employee Performance scoring or dashboard content.

## Data and calculations

- It loads live from `/api/employee-capacity-utilization/data` whenever the page opens. The endpoint reads the server's canonical database, using the active successful `canonical_issues` snapshot and newest persisted `canonical_worklogs` run. The deployed HTML does not rely on build-time database contents.
- Capacity comes from the selected capacity profile. Weekday official leaves in that profile are excluded from capacity before Availability is calculated; Availability is then capacity less planned and unplanned leave. Weekend official-leave dates remain visible in the Official Leaves count but do not reduce hours because they are already non-working days.
- Booked Manhours is the original estimate of assigned subtasks overlapping the selected month. Logged Hours has two clearly labelled scopes: **All work logged by employee** counts every in-month Jira worklog authored by that employee; **Work logged on their assigned subtasks** counts only time on subtasks currently assigned to the same employee. The narrower scope excludes stories, tasks, bugs, and work items assigned to somebody else. RLT-board leave worklogs are excluded by default because leave hours are already reported in Leaves Taken; the user can explicitly include them when needed.
- The final Grand Total row sums the displayed employees. Official Leaves shows the selected profile's holiday-day count per employee and the corresponding resource-day total in the Grand Total row.

## Filters

The filter bar is divided into two purposeful rows. **Refine the report** contains the primary scope choices—Month, Capacity Profile, Teams, and the prominently highlighted Logged Hours source. **Display & inclusion** contains presentation and visibility preferences—color formatting, leave-worklog inclusion, resigned employees, and Support badges. All changes apply immediately; the status pill confirms this behaviour. Use **Hide controls** to collapse the panel: its persistent Active filters chips retain the selected month, profile, teams, logged-hours scope, and enabled display options. The choice persists in browser storage. Team and color menus open above the panel rather than being clipped by it.

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

Availability opens **Availability Details**, a centered modal occupying 80% of desktop viewport width (96% on narrow screens), above the navigation and report. Its full-month Monday-first calendar uses equal-size cells, including empty leading/trailing cells. Grey `0h logged` means no subtask work; green hours are clickable subtask time; red hours are planned plus unplanned leave taken. Both badges appear on mixed work/leave days. Weekend and official-leave labels provide context without hiding logged work.

The calendar always counts non-RLT subtask worklogs authored by the selected employee, including subtasks assigned to someone else. It is independent of the summary Logged Hours scope and Include leaves switch; it therefore need not match the summary Logged Hours value. Grand Total Availability aggregates the displayed employees. Summary cards show available capacity, subtask hours, and leave taken. Select green hours to display the day's subtasks below, grouped by issue and author, with Jira links, titles, author, Start/Due, original estimates, and summed hours. Start/Due show the date precision available in canonical data; missing dates show a dash. Original estimates appear once per grouped subtask, rather than being multiplied by worklog count.

Other metric drawers can be widened or narrowed by dragging their left resize handle; that width persists in browser storage. Availability retains its calendar width. Dialogs close through their close button, backdrop, or Escape, trap keyboard focus, and return focus to the triggering table cell. Team and color dropdowns render above the sticky table header.

## Excel export

Download Excel exports the complete dataset for the active Month, Capacity Profile, Teams, Logged Hours scope, leave-worklog inclusion, and resigned-resource setting. The workbook contains Export Info, Summary, Worklogs, Booked Subtasks, Leave Records, Capacity Calendar, and Employees worksheets. Work item and Epic URLs are active Excel hyperlinks. Each data sheet uses frozen headers, filters, table styling, readable column widths, and numeric hour/utilization columns.

## Related code

- `generate_employee_capacity_utilization_report.py`
- `employee_capacity_utilization_export.py`
- `report_server.py`
- `run_html_only.py`
- `shared-nav.js`
- `employee_availability_calendar.js` and `.css` — calendar model, daily drilldown, modal styling, and dropdown layering; inlined by the generator into the report so packaged reports need no additional asset requests.
- `tests/test_employee_availability_calendar.cjs` and `tests/employee-availability.playwright.config.cjs` — focused browser regression coverage.

## Business Cases

Managers can inspect a person's full month before scheduling work, distinguish missing subtask logs from personal leave, and trace a day's hours back to Jira.

## Examples

Two Alice worklogs of 2h and 3.5h on APP-1 on October 5 show one green 5.5h badge and one detail row with 5.5h. If Alice also took 3h leave, a red 3h badge appears in the same cell. Bob's logs, stories, RLT worklogs, and September logs do not contribute to Alice's October subtask hours.

## Explanations

Choose a month, click an employee's Availability value, scan the colored daily badges, and click green hours to see the underlying subtasks. Leave stays visible even when no work was logged.

## Table Schema

This change adds no tables or columns. The read-only payload derives from `canonical_issues` (run, issue key/type, assignee, summary, hierarchy, start/due dates, original estimate), `canonical_worklogs` (run, issue, author, started date, hours), and the existing RLT daily leave snapshot (`assignee`, `period_day`, `planned_taken_hours`, `unplanned_taken_hours`). Capacity profiles, configured teams, support members, and resource resignation settings retain their existing readers and constraints.

## Dependent & Impacted Files

`report_server.py` continues to serve the existing live-data API and promote root report HTML. `run_html_only.py` and `run_server.py` invoke the generator. `employee_capacity_utilization_export.py` retains the existing summary/export scope semantics; the calendar does not change exported calculations. The report overview links this screen's behavior.

## Data Flow

The server builds the canonical payload on page load. The generator inlines the calendar JavaScript and CSS into the root report shell, which is copied/promoted to `report_html`. Availability joins authored worklogs to issue metadata, filters non-RLT subtasks by employee and month, groups them by day/issue/author, and joins daily leave taken. Selecting a green badge renders grouped source rows locally without another API call.

## Local verification

Availability values are bold purple and underlined. Closing the calendar hides the complete modal without leaving a strip at the right edge. Each in-month day has a hover/focus plus button for reviewed Jira work logging. Personal OAuth setup, confirmation steps, author checks, request receipts, snapshot timing and dependencies are documented in [Calendar worklogs and Jira sign-in](../../../JIRA_WORKLOG_SIGNIN.md).

The calendar now uses compact 102px equal-height cells, 6px grid gaps, and reduced panel/card padding (about 20% less vertical space). Today uses the viewer's local calendar date, a blue date badge and border, a visible Today label, and `aria-current="date"`. It remains highlighted when another day is selected, and no Today marker appears when viewing another month.

Red leave badges are buttons: clicking or pressing Enter opens the selected day's leave source table below the calendar. Rows show employee, date, all associated Jira leave links, planned taken hours, unplanned taken hours, and total taken hours. The existing daily leave snapshot aggregates by employee/day, so multiple associated Jira links share that daily total; hours are not duplicated per link. Selecting work or leave replaces the same detail section and updates the selected badge. These rows use the same taken-hour data as the red badge, independent of the Logged Hours scope and Include leaves toggle. No API or database schema changes are needed.

Run `.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py`, then `.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3118`. Open `http://127.0.0.1:3118/employee_capacity_utilization_report.html`. Startup promotes the root shell. Tests: `.venv/Scripts/python.exe -m unittest tests.test_employee_capacity_utilization_report` and `node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line`.

## Change notes

- The report now uses production runtime data instead of the canonical payload embedded when the HTML was generated.
- Filters were reorganized into a responsive control panel with switch controls, a team-count summary, and a sticky, readable data table.
- The Teams dropdown uses compact aligned checkboxes, single-line team labels, sticky selection actions, a bounded scroll area, and closes on outside click or Escape.
- Utilization thresholds and row-versus-column color scope are user-configurable, validated as `0 ≤ A < B < C ≤ 500`, and saved per browser.
- Every summarized value opens a resizable, metric-specific raw-data drawer.
- Employee Name opens hour logs, worklogs include Epic context, and the current filtered dataset can be downloaded as a structured multi-sheet Excel workbook.
- RLT-board leave worklogs are excluded from Logged Hours by default and can be restored instantly with Include leaves in logged hours; the same selection is applied to the table, drawer, totals, utilization, colors, and Excel export.
- Official-leave calendar calculations use local noon for date matching, preventing UTC date shifts from causing weekday official leave to remain in availability for users east of UTC.
