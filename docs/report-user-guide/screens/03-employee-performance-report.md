# Employee Performance Report

Report ID: `employee_performance`

INFO_IDS: `employee.team_avg_score`, `employee.advanced_score_sum`, `employee.capacity_per_employee`, `employee.planned_hours_assigned`, `employee.assigned_counts`, `employee.missed_start_ratio`

## Key Fields

| Field | Definition | Formula | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Team Avg Score | Mean weighted assignee performance score. | `Average(sum(effective_weight * factor_score_pct / 100))` | normalized factor scores, configured weights | Advanced Score weights from performance settings must total 100. | Assignee workload and leave impact context. |
| Advanced Score Sum | Sum of eligible assignee Advanced Scores in the current filter. | `Sum(assignee_weighted_advanced_score)` | six normalized Advanced Score factors | Settings from `/api/performance/settings` must be valid. | Dashboard risk indicators and missed planning quality. |
| Employee Capacity | Effective per-assignee capacity after leave adjustment. | `baseline_capacity - planned_leave_hours - unplanned_leave_hours` | capacity profile, leave hours | Falls back to weekday `*` 8h if profile missing. | Capacity subtraction and nested capacity gap. |
| Planned Hours Assigned | Estimate hours assigned to employee in selected window. | `Sum(original_estimate_hours over assigned items)` | estimates, assignee, date range | Missing estimates count as 0h. | Planned-vs-actual workload context. |
| Assigned Item Counts | Hierarchy split of assigned work items. | `Count(Epics, Stories, Subtasks)` | issue type, assignee, parent linkage | Unknown types are mapped to subtask bucket. | Dashboard hierarchy density and delivery mix. |
| Missed Start Ratio | Late-start ratio over assigned items. | `missed_start_count / total_assigned_count` | planned start date, worklog date, assignee | Start-day context uses planned start date only. | Missed entries and employee risk trend. |
| Simple Score | Estimate-compliance score shown as the primary assignee score. | `clamp(100 * (1 - adjusted_overrun / total_estimated_hours), 0, 100)` | original estimates, actual logged hours, commitment overruns | With `Include Due Completion` on, over-estimate subtasks completed on or before due date forgive their overrun hours. | Assignee drilldown Simple Scoring section, leaderboard simple mode, and report-level Simple Scores toggle. |
| Advanced Score (beta) | Weighted normalized score shown beside the Simple Score for comparison. | `Sum(effective_weight * factor_score_pct / 100)` | Estimate Discipline, Due-Date Delivery, Subtask Timeliness, Bug Quality, Late-Bug Severity, Leave Reliability | The six configured weights must total 100; factors with a zero denominator are N/A and their weight is redistributed. | Assignee drilldown Advanced Scoring section and leaderboard advanced mode. |

## Drawer Notes

- The assignee drilldown shows both Simple Scoring and Advanced Scoring sections. Simple Scoring includes the `Include Due Completion` toggle, a per-subtask estimate-vs-actual table, and a compliance donut. Advanced Filters, Leaderboard controls, and Executive Scorecards performance controls expose synchronized `Simple Scores` / `Advanced Scores` toggles for the report-level score display and ranking mode. Advanced Scoring now shows a factor breakdown with input, denominator, effective/configured weight, factor %, and contribution.
- The Simple Score details drawer explains applied overrun, commitment forgiveness, overload handling, and the exact formula inputs for the selected assignee. The Advanced Score details drawer uses the same drawer modal to show normalized weighted factor contributors, final advanced score, and the same All Scored Subtasks table with Epic/RMI and Project filters.
