# Dashboard Report

Report ID: `dashboard`

INFO_IDS: `dashboard.delivery_health`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Delivery Health Cards | Card view of execution status, schedule, effort, and IPP sync state. | Hierarchy rollup over epics/stories/subtasks with date/status filters. | status, planned dates, actual dates, logged hours, IPP flags | Hierarchy integrity, orphan handling, active filters. | Missed Entries quality, Employee Performance penalties, IPP roadmap. |

## Drawer Notes

- `i` icon on lane sections opens structured explanation for how delivery health cards are built and interpreted.
