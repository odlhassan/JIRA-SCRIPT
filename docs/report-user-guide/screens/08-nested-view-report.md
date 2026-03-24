# Nested View Report

Report ID: `nested_view`

INFO_IDS: `nested.capacity_gap`, `nested.total_capacity_adjusted`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Availability for more work | Capacity remaining after planned project load and planned leave estimates. | `Total Capacity - Total Planned Projects - Total Leaves Planned` | total capacity, planned projects (non-RLT), leaves planned (RLT) | Date range and project filter sensitive; RLT rules applied. | Assignee subtraction, RnD investable hours. |
| Availability | Practical capacity after planned leaves are deducted. | `Total Capacity (Hours) - Total Leaves Planned` | total capacity, planned leaves (RLT original estimates) | Date range and project filter sensitive; RLT leave planned estimates are deducted once. | RLT leave totals and RnD leave-adjusted capacity. |

## Drawer Notes

- Drawer uses live breakdown terms to explain profile capacity, leaves, and downstream planning impact.
- On load, the first available saved capacity profile is auto-applied to KPI calculations.
- On desktop, the drawer opens at 50% viewport width and can be resized by dragging its left edge.
- The Capacity Profile drawer now includes a read-only calendar preview for the currently selected profile.
- The preview follows the active Nested View date filter and shows:
  - summary chips for range capacity, business days, holiday weekdays, per-assignee capacity, and leave totals
  - month cards with Ramadan (`R`), holidays (`H`), leave (`L`), and Ramadan leave (`RL`) tags
- `Total Capacity (Hours)` uses: `Employee Count x Available Business Days x Per Day Hours`, and shows this as an icon-based top-right formula chip.
- `Total Leaves Taken` is currently hidden/commented in the scorecard UI (reserved for later activation).
