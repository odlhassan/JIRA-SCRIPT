# Team RMI Gantt Report

Report ID: `phase_rmi_gantt`

INFO_IDS: `phase_rmi.weekly_load`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Weekly Load Chips | Weekly team-lane workload intensity summary. | Aggregate epic man-days (story estimate hours / 8) over overlapping week buckets. | story assignee team mapping, story parent epic, story estimate hours, epic start/end | Only stories with parent epic, valid date range, and positive estimate participate in weekly load. | Team timeline spread and roadmap workload pressure. |
| Team Lane | Team-level row showing all RMIs on that team's plate in selected date range. | Group by team inferred from story assignee mapped in `performance_teams`. | story assignee, `performance_teams.assignees_json` | Assignees not mapped to any team appear under `Unmapped Team`. | Team planning and ownership alignment. |
| RMI/Epic Cards | Clickable RMI cards in each team lane. | Aggregate stories by `(team, parent epic)` to compute total planned hours/man-days and date span. | story rows, epic metadata, Jira URL fallback | Cards must render Jira link and open epic in a new tab. | Drill-through from workload view to Jira execution. |

## Drawer Notes

- Drawer explains weekly load chips for team-level planning risk.
- Data source is SQLite snapshot (`team_rmi_gantt_items`) built by `sync_team_rmi_gantt_sqlite.py`.
