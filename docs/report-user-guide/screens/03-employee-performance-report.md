# Employee Performance Report

Report ID: `employee_performance`

INFO_IDS: `employee.team_avg_score`, `employee.total_penalty`, `employee.capacity_per_employee`, `employee.planned_hours_assigned`, `employee.assigned_counts`, `employee.missed_start_ratio`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Team Avg Score | Mean assignee performance score. | `Average(clamp(base_score - total_penalty, min_score, max_score))` | base score, penalty totals, clamp bounds | Penalty multipliers from performance settings. | Assignee workload and leave impact context. |
| Total Penalty | Sum of weighted penalty contributors. | `Sum(weighted penalties across bug, delay, overrun, unplanned leave)` | bug hours, late hours, overrun, leave, settings weights | Settings from `/api/performance/settings` must be valid. | Dashboard risk indicators and missed planning quality. |
| Employee Capacity | Effective per-assignee capacity after leave adjustment. | `baseline_capacity - planned_leave_hours - unplanned_leave_hours` | capacity profile, leave hours | Falls back to weekday `*` 8h if profile missing. | Capacity subtraction and nested capacity gap. |
| Planned Hours Assigned | Estimate hours assigned to employee in selected window. | `Sum(original_estimate_hours over assigned items)` | estimates, assignee, date range | Missing estimates count as 0h. | Planned-vs-actual workload context. |
| Assigned Item Counts | Hierarchy split of assigned work items. | `Count(Epics, Stories, Subtasks)` | issue type, assignee, parent linkage | Unknown types are mapped to subtask bucket. | Dashboard hierarchy density and delivery mix. |
| Missed Start Ratio | Late-start ratio over assigned items. | `missed_start_count / total_assigned_count` | planned start date, worklog date, assignee | Start-day context uses planned start date only. | Missed entries and employee risk trend. |

## Drawer Notes

- Drawer clarifies each penalty source and the leadership impact on score trend interpretation.
