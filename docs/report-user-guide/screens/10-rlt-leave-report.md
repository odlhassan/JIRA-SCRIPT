# RLT Leave Report

Report ID: `rlt_leave_report`

INFO_IDS: `rlt.total_taken`, `rlt.total_planned_leaves`, `rlt.future_planned`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Total Taken | Leave already consumed in selected range. | `Planned Taken + Unplanned Taken` | planned taken hours, unplanned taken hours | Days are derived from configured daily-hours logic. | Nested leave-adjusted capacity and RnD capacity baseline. |
| Total Leaves Planned | Total planned leave load in selected range (taken + not-yet-taken). | `Planned Taken + Future Planned` | planned taken hours, planned not-yet-taken hours | Uses the same filtered range as RLT scorecards. | Nested total capacity adjusted and capacity gap. |
| Future Planned | Planned leave not yet consumed. | `Sum(planned estimates not yet taken)` | planned not-yet-taken hours | Missing required leave metadata is tracked in No Entry. | Nested adjusted capacity and assignee leave totals. |

## Drawer Notes

- Drawer describes leave categories and how they are reused in capacity-linked reports.
