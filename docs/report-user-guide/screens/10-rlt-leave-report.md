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

## Verification Signals

- The report uses `canonical_issues` and `canonical_worklogs` from `assignee_hours_capacity.db` by default, scoped to the last successful canonical refresh run and the RLT project. Direct Jira loading is a legacy, explicit `--source jira` option.
- The Verification Signals table and Excel sheet include both `start_date` and `due_date` alongside the created timestamp and derived verification reference date, so reviewers can compare the source leave window with the late-creation signal.
