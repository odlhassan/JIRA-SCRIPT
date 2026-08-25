# Employee Capacity & Utilization

## Purpose

This standalone report presents employee capacity, profile holidays, personal leave, booked subtask estimates, logged hours, and utilization without Employee Performance scoring or dashboard content.

## Data and calculations

- It requires a successful canonical refresh and reads `canonical_issues` plus the newest persisted `canonical_worklogs` run.
- Capacity comes from the selected capacity profile. Availability is capacity less planned and unplanned leave.
- Booked Manhours is the original estimate of assigned subtasks overlapping the selected month. Logged Hours can use all employee worklogs or only worklogs against their assigned subtasks.
- The final Grand Total row sums the displayed employees. Official Leaves is the profile holiday-day count, shown once.

## Filters

| Field | Default | Effect |
| --- | --- | --- |
| Month | Current month | Limits worklogs, leave, and booked-subtask date overlap. |
| Capacity profile | Auto | Selects the capacity calendar. |
| Teams | All except Process Team | Limits employees to selected configured teams. |
| Logged hours | Any employee worklog | Switches all worklogs vs assigned-subtask worklogs. |
| Display resigned | Off | Includes resigned resources. |
| Indicate support | On | Adds a Support chip for `support_team_config` members. |

## Related code

- `generate_employee_capacity_utilization_report.py`
- `report_server.py`
- `run_html_only.py`
- `shared-nav.js`
