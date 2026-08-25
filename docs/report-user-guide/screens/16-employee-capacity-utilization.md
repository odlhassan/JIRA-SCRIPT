# Employee Capacity & Utilization

## Purpose

This standalone report presents employee capacity, profile holidays, personal leave, booked subtask estimates, logged hours, and utilization without Employee Performance scoring or dashboard content.

## Data and calculations

- It loads live from `/api/employee-capacity-utilization/data` whenever the page opens. The endpoint reads the server's canonical database, using the active successful `canonical_issues` snapshot and newest persisted `canonical_worklogs` run. The deployed HTML does not rely on build-time database contents.
- Capacity comes from the selected capacity profile. Availability is capacity less planned and unplanned leave.
- Booked Manhours is the original estimate of assigned subtasks overlapping the selected month. Logged Hours can use all employee worklogs or only worklogs against their assigned subtasks.
- The final Grand Total row sums the displayed employees. Official Leaves shows the selected profile's holiday-day count per employee and the corresponding resource-day total in the Grand Total row.

## Filters

| Field | Default | Effect |
| --- | --- | --- |
| Month | Current month | Limits worklogs, leave, and booked-subtask date overlap. |
| Capacity profile | Auto | Selects the capacity calendar. |
| Teams | All except Process Team | Includes unassigned employees and employees in selected configured teams; Process-Team-only resources are hidden by default. Changes apply immediately. |
| Logged hours | Any employee worklog | Switches all worklogs vs assigned-subtask worklogs. |
| Display resigned | Off | Includes resigned resources. |
| Indicate support | On | Adds a Support chip for `support_team_config` members. |

There is no Apply or Refresh button. Every filter change redraws the table immediately. If canonical data cannot be loaded, the report shows an explicit error rather than a misleading table of zero hours.

## Related code

- `generate_employee_capacity_utilization_report.py`
- `report_server.py`
- `run_html_only.py`
- `shared-nav.js`

## Change notes

- The report now uses production runtime data instead of the canonical payload embedded when the HTML was generated.
- Filters were reorganized into a responsive control panel with switch controls, a team-count summary, and a sticky, readable data table.
- The Teams dropdown uses compact aligned checkboxes, single-line team labels, sticky selection actions, a bounded scroll area, and closes on outside click or Escape.
