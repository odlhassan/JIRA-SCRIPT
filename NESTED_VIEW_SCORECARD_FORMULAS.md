# Nested View Scorecards Field Definitions

This document captures the scorecard field definitions and formulas used in `nested_view_report.html` and mirrored in `assignee_hours_report.html`.

Cross-report drawer and field-logic documentation now also lives in:

- `docs/report-user-guide/00-report-overview.md`
- `docs/report-user-guide/screens/08-nested-view-report.md`

## Scope
- Date range: user-selected From/To range on the report.
- Leave project: `RLT` (`RLT RnD Leave Tracker`).

## Fields and Formulas

| Field Caption | Definition | Formula |
|---|---|---|
| Total Capacity (Hours) | Capacity computed from applied profile factors for selected range (fallback: project totals). | `Total Capacity (Hours) = Employee Count x Available Business Days x Per Day Hours` |
| Total Leaves Planned | Planned leave load from day-bucketed RLT rows in selected date range. | `Total Leaves Planned = Planned Taken + Planned Not Yet Taken` |
| Availability | Capacity reduced only by planned leaves from RLT. | `Availability = Total Capacity (Hours) - Total Leaves Planned` |
| Capacity available for more work | Capacity after subtracting planned project work and planned RLT leave load. | `Capacity available for more work = Total Capacity (Hours) - Total Planned Projects (Hours) - Total Leaves Planned` |

## Notes
- `Total Planned Projects (Hours)` excludes `RLT`.
- `Total Actual Project Hours` excludes `RLT`.
- First available capacity profile is auto-applied on Nested View load (when profiles exist).
- `Total Leaves Taken` scorecard is currently commented out in Nested View UI and retained for future use.
- Assignee Hours Report now displays these Nested View scorecards in a dedicated section, in addition to existing Assignee scorecards, without duplicating card instances.
