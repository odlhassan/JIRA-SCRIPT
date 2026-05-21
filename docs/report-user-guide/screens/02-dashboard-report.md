# Dashboard Report

Report ID: `dashboard`

INFO_IDS: `dashboard.delivery_health`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Delivery Health Cards | Card view of execution status, schedule, effort, and IPP sync state. | Hierarchy rollup over epics/stories/subtasks with date/status filters. | status, planned dates, actual dates, logged hours, IPP flags | Hierarchy integrity, orphan handling, active filters. | Missed Entries quality, Employee Performance penalties, IPP roadmap. |

## Drawer Notes

- `i` icon on lane sections opens structured explanation for how delivery health cards are built and interpreted.

## Shared Refresh Widget Dependency

- Dashboard-style reports, including Employee Performance, use the shared refresh widget injected by `report_server.py`.
- That shared widget includes the `rw-busy-modal` overlay shown only when a new refresh request collides with an already-running refresh.
- Regressions in the shared busy-modal CSS can affect both the dashboard family and Employee Performance layouts, so linked screen docs track that dependency explicitly.

## Dependent & Impacted Files

- `report_server.py` provides the shared refresh button, progress UI, and busy-lock modal for dashboard-style reports.
- `docs/report-user-guide/screens/03-employee-performance-report.md` documents the Employee Performance-specific usage of the same shared refresh-lock overlay.
