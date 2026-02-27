# RnD Data Story Report

Report ID: `rnd_data_story`

INFO_IDS: `rnd.leave_adjusted_capacity`, `rnd.pending_hours_required`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Leave-Adjusted Capacity | Capacity after all leave impact deductions. | `Available Capacity - (Planned Taken + Planned Not Taken Yet + Unplanned Taken)` | available capacity, planned taken, planned not yet taken, unplanned taken | RnD scope fixed; valid date filter required. | Nested adjusted capacity and RLT leave totals. |
| Pending Hours Required | Remaining effort required for scoped epics. | `Sum(max(Epic Estimate - Epic Logged, 0))` | epic original estimates, epic logged hours | Epic date inclusion uses start OR end in range; no negatives. | Assignee plan-actual gap and Nested completion hours. |

## Drawer Notes

- Drawer supports leadership decisions on commitment coverage and demand-risk gap.
