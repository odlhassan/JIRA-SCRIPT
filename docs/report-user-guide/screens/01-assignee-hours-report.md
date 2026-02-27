# Assignee Hours Report

Report ID: `assignee_hours`

INFO_IDS: `assignee.capacity_subtraction`, `assignee.project_plan_actual_gap`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Capacity Subtraction (Hours) | Remaining capacity after actual work and leave impact. | `Available Capacity - Project Actual Hours - Leave Hours` | available capacity, project actual (non-RLT), leave total | Valid date range; RLT excluded from project actual. | Nested capacity gap, RnD investable hours. |
| Project Plan - Actual Hours | Remaining planned work after logged work. | `Project Planned Hours - Project Actual Hours` | project planned (non-RLT), project actual (non-RLT) | Uses active report filters. | Nested hours required, RnD pending demand. |

## Drawer Notes

- `i` buttons open the right drawer.
- Drawer includes definition, formula, ingredients, validations, in-report links, cross-report links, and data provenance.
