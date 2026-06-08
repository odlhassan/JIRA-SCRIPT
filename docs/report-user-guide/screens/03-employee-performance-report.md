# Employee Performance Report

Report ID: `employee_performance`

INFO_IDS: `employee.team_avg_score`, `employee.advanced_score_sum`, `employee.capacity_per_employee`, `employee.planned_hours_assigned`, `employee.assigned_counts`, `employee.missed_start_ratio`

## Business Logic

- Date scope defaults to the current month. `Epics By TK Dates` includes epics whose Epic or Story/TK start and due dates are fully inside the selected range. `Epics By Subtask Dates` includes subtasks whose planned start and due dates are fully inside the selected range.
- Team filtering is team-level: selected teams expand to their configured members from `performance_teams.assignees_json`, and only matching assignees remain in the computed leaderboard and KPI set.
- The Teams filter uses a custom menu rather than a browser-native dropdown. It groups members under their team, shows the team leader and member count, and decorates members only when extra context is useful: `Resigned` for resigned resources and `Support team` for support-roster members.
- Resignation status is read from `performance_resource_resignations`: a row means `Resigned` and may include a resignation date; active/non-resigned members intentionally show no badge to keep the menu clean.
- `Support team` is read read-only from `support_team_config.members_json`; the report does not create or modify that table.
- Employee capacity is `base_capacity_hours - planned_leave_hours`. Unplanned leave is displayed but not deducted from capacity.
- Simple Score is `clamp(100 * (1 - adjusted_overrun / total_estimated_hours), 0, 100)`. When due-completion is enabled, over-estimate subtasks finished on or before due date forgive their overrun hours.
- Advanced Score is the weighted normalized sum of Estimate Discipline, Due-Date Delivery, Subtask Timeliness, Bug Quality, Late-Bug Severity, and Leave Reliability. Weights must total 100; factors with zero denominator are N/A and their weight is redistributed.

## Business Cases

- Engineering leads use the report to compare employee delivery performance, capacity pressure, overburn, and planning discipline for a selected time window.
- Team leads use the custom Teams filter to filter by organization team while seeing the exact members included before applying the filter.
- Management can distinguish active, resigned, and support-team resources directly inside the filter, reducing mistakes when reviewing cross-team performance or support capacity contributors.

## Examples

| Input | Output |
| --- | --- |
| Team `Fullstack Team`, leader `Ameer Hamza Khan`, members `Ameer Hamza Khan`, `Sarmad Sabir`; `Sarmad Sabir` in `support_team_config`; no resignation row | Teams menu shows a `Fullstack Team` card, both members nested below it, and `Sarmad Sabir` with only a `Support team` chip. |
| Member `Maria Sharafat` appears in `performance_resource_resignations` with `2026-04-06` | Member row shows `Resigned 6 Apr 2026`. |
| Date range `1 Mar` to `31 Mar`, `Epics By TK Dates`, selected project `O2` and selected team `SQA` | Only SQA members with qualifying O2 work in March contribute to scorecards, leaderboard, detail drawer, and scoped subtask actuals. |

## Explanations

The generator loads work items, worklogs, leave rows, team definitions, resignation records, support-team membership, capacity profiles, settings, and simple-scoring subtasks. It embeds these as JSON in `employee_performance_report.html`. The browser computes per-assignee metrics for the current filters, renders executive scorecards, the leaderboard, team chart, and the assignee drilldown.

The Teams filter is still backed by a hidden multi-select (`#teams`) so existing filtering code remains stable. The visible control is a custom popover: each configured team is a spacious card with an `Include team` checkbox and nested member rows. Search matches both team names and member names. Reset selects all teams and resynchronizes the custom visual state.

## Front-end UI Fields

| Field | Type | Default | Valid values / range | Controls | Example |
| --- | --- | --- | --- | --- | --- |
| Filters | button + popover | closed | open/closed | Opens date, project, team, and epic-basis filters | Click `Filters` to reveal the custom Teams menu. |
| From / To | date inputs | current month start/end | ISO date | Date window for worklogs, leave rows, and qualifying epics/subtasks | `2026-03-01` to `2026-03-31`. |
| Date presets | buttons | none | Last 30 Days, Last Month, Current Month, Last 90 Days, Last Quarter, Current Quarter | Quickly writes From/To | `Current Month` fills the current month. |
| Project | custom multi-select | all projects selected | generated project keys, displayed with managed names | Limits work and scoped subtasks to selected projects | `O2` shown as configured display name. |
| Team | custom grouped multi-select | all teams selected | teams from `performance_teams` | Limits assignees to members of selected teams | Select only `SQA`. |
| Team search | text input | blank | free text | Filters team cards by team name, leader, or member | Type `Abbas` to find teams containing Muhammad Abbas. |
| Include team | checkbox | checked | checked/unchecked | Selects or clears that team in the hidden `#teams` select | Clear `Fullstack Team` to remove its members. |
| Member chips | labels | computed | `Resigned [date]`, `Support team` | Display-only context for nested members only when a badge adds meaning | `Resigned 6 Apr 2026` + `Support team`. |
| Epics By | segmented buttons | TK Dates | TK Dates, Subtask Dates | Chooses date scoping basis | `Subtask Dates` uses subtask planned dates. |
| Settings | button + popover | closed | open/closed | Capacity profile, overburn basis, efficiency mode, score display | Switch to Advanced Scores. |
| Capacity Profile | select | Auto | saved capacity profiles | Chooses capacity calendar for employee capacity | `Auto (Match selected date range)`. |
| Overburn | select/buttons | Overburn Per Task | Per Task, Total | Changes Simple Score overrun basis | Total penalizes only aggregate overrun. |
| Efficiency | select/buttons | Penalty Inclusive Efficiency | Penalty Inclusive, Simple | Changes executive efficiency card | Simple compares total planned vs actual. |
| Score Display | segmented buttons | Simple Scores | Simple, Advanced | Changes headline average and leaderboard score mode | Advanced ranks by weighted normalized score. |
| Search Assignee / Leaderboard Search | text inputs | blank | free text | Filters detail or leaderboard by assignee name | Search `Sarmad`. |
| Leaderboard sort / direction | selects | Performance Score / Descending | configured sort modes / asc-desc | Sorts leaderboard rows | Sort by Capacity Gap ascending. |
| At-Risk View | select | All Assignees | All, Only At-Risk | Filters scores below 60 | Show only score `<60`. |
| Start Discipline | select | All | All, Only Missed Starts | Filters assignees with missed planned starts | Show only missed starts. |
| Drilldown toggles | checkboxes/buttons | off/default | on/off | Extended actuals, overload penalty, due completion, drawer sections | Enable Extended Actuals to recompute actuals. |

## Script Files

| File | Role |
| --- | --- |
| `generate_employee_performance_report.py` | Canonical generator, SQLite readers, scoring precompute, embedded CSS/JS, custom Teams menu. |
| `employee_performance_report.html` | Generated root report served by localhost after sync. |
| `report_html/employee_performance_report.html` | Served copy used by `/employee_performance_report.html`. |
| `report_server.py` | Injects shared refresh widget and serves APIs used by the report (`/api/performance/settings`, `/api/scoped-subtasks`). |
| `tests/test_employee_performance_report.py` | Focused generator, payload, scoring, and HTML contract tests. |
| `docs/report-user-guide/screens/03-employee-performance-report.md` | Module behavior and UI documentation. |

## Dependent & Impacted Files

- `report_server.py` injects the shared refresh widget, busy-modal CSS, and 409-conflict handling used by this screen.
- `monthly_epic_plan_progress_service.py` owns the `support_team_config` table and support-team roster model that the Employee Performance Teams filter reads read-only for member chips.
- `support_center_service.py` and `SUPPORT_CENTER_REPORT.md` also depend on the same support-team roster for support availability reporting.
- `tests/test_report_date_filter_api.py` verifies the served Employee Performance HTML keeps valid busy-modal overlay CSS.
- `docs/report-user-guide/screens/02-dashboard-report.md` tracks the same shared refresh widget behavior used by dashboard-style reports.

## Table Schema

### `performance_teams`

| Column | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| `team_name` | TEXT | PRIMARY KEY | Display name and selected value for each team. |
| `team_leader` | TEXT | NOT NULL DEFAULT '' | Team lead shown in the Teams menu and team performance card. |
| `assignees_json` | TEXT | NOT NULL | JSON array of member names used for team filtering. |
| `updated_at` | TEXT | NOT NULL | Last saved timestamp. |

### `performance_resource_resignations`

| Column | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| `assignee_name` | TEXT | PRIMARY KEY | Employee name. Row existence means resigned. |
| `resignation_date` | TEXT | nullable | Optional ISO date shown in the Teams menu. |
| `updated_at` | TEXT | NOT NULL | Last saved timestamp. |

### `support_team_config` (read-only in this report)

| Column | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| `key` | TEXT | PRIMARY KEY | `members` row stores the support roster. |
| `members_json` | TEXT | NOT NULL DEFAULT '[]' | JSON array of support-team member names. |
| `updated_at` | TEXT | nullable/defaulted by owner module | Last saved timestamp. |

### Other tables read

- `performance_point_settings` stores Advanced Score weights and overload/planning-realism switches.
- `simple_scoring_subtasks` stores precomputed estimate-vs-actual rows for Simple Score.
- `canonical_issues`, `canonical_worklogs`, `canonical_refresh_state`, and EPF snapshot tables provide Jira issue/worklog source data when DB-backed modes are active.

## Data Flow

1. `main()` resolves source mode (`auto`, `db`, `canonical_db`, or `xlsx`) and loads settings, teams, support-team members, capacity profiles, work items, worklogs, and leave rows.
2. `_precompute_simple_scoring()` persists subtask estimate/actual/completion details to `simple_scoring_subtasks`.
3. `_load_performance_resource_resignation_map()` reads resignation records for configured team members, and `_load_support_team_members()` reads `support_team_config` without creating tables.
4. `_build_payload()` embeds teams, `resource_records`, and `support_team_members` into the report JSON.
5. `_build_html()` emits the custom Teams filter. The hidden `#teams` select stores selected team names; `setupTeamFilterDropdown()` renders team cards and member chips and keeps the select synchronized.
6. Browser `renderAll()` loads scoped subtasks for current filters, computes assignee metrics, applies team/project/date filters, and renders scorecards, leaderboard, charts, and drilldowns.

## Drawer Notes

- The assignee drilldown shows both Simple Scoring and Advanced Scoring sections. Simple Scoring includes the `Include Due Completion` toggle, a per-subtask estimate-vs-actual table, and a compliance donut.
- The Simple Score details drawer explains applied overrun, commitment forgiveness, overload handling, and the exact formula inputs for the selected assignee.
- The Advanced Score details drawer uses the same drawer modal to show normalized weighted factor contributors, final advanced score, and the All Scored Subtasks table with Epic/RMI and Project filters.

## Refresh Lock UX

- The `Refresh` control on this screen uses the shared `report_server.py` refresh widget rather than inline page-specific markup.
- When another refresh run is already active, the server returns a busy response and the page opens the `rw-busy-modal` overlay with progress, elapsed time, estimated remaining time, current step, and report name.
- The overlay container must remain hidden by default with `.rw-busy-overlay { display: none; position: fixed; inset: 0; }` and only switch to visible when JavaScript sets `aria-hidden="false"` after a conflicting refresh attempt.
