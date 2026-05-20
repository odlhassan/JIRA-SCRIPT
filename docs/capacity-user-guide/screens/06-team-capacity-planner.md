# Team Capacity Planner

## Screen

- Name: Team Capacity Planner
- Route: `/settings/team-capacity-planner`
- Purpose: Review team capacity, leave impact, logged work, assigned epics, and per-resource planned subtask load.

## Sections

### Toolbar

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Team (`team-select`) | Dropdown | Yes | First configured performance team | Chooses the resource group to display. | Values come from `performance_teams`. |
| From (`from-date`) | Date | Yes | First day of current month | Start of the capacity and planned-work window. | Must be a valid ISO date. |
| To (`to-date`) | Date | Yes | Last day of current month | End of the capacity and planned-work window. | Must be a valid ISO date. |
| Capacity Profile (`profile-select`) | Dropdown | No | Auto best match | Selects the working-day profile used to convert hours to capacity/day stats. | Auto mode picks the best matching saved profile. |
| Stats (`stat-unit-toggle`) | Segmented toggle | No | Days | Switches resource stats between days and hours. | Toggle is display-only; it does not reload or change saved data. |
| Analytics (`analytics-btn`) | Icon + label button | No | — | Opens the Team Analytics right-side drawer showing aggregate team stats. | Amber-orange gradient button. Active only after team data is loaded. Respects current stat unit (Days/Hours). |
| Load (`load-btn`) | Button | Yes | — | Fetches team capacity data for the selected team and date range. | Disabled while loading. |

### Team Members

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Planned | Bar/value | No | `0d` | Planned load for the resource. | Sums only assigned subtask original estimates in the selected date range. Assigned epic and story estimates are intentionally ignored. |
| Logged | Bar/value | No | `0d` | Logged work for the resource in the selected date range. | Uses canonical worklogs by issue assignee. |
| Available | Bar/value | No | Derived | Remaining capacity after leave impact. | Uses selected/auto capacity profile and leave rows. |
| Assigned Epics | Expandable list | No | Empty | Epics associated with the resource's assigned work items. | Used for navigation/highlighting only; not used directly in the Planned value. |

### Team Analytics Drawer

Opened by clicking the **Analytics** button in the toolbar. Slides in from the right with a dark backdrop. Scrollable — all sections are visible via vertical scroll.

| Section | Description |
| --- | --- |
| Team pill | Team name, member count, and selected date range. |
| Total Capacity | Gross working time (sum of all members' `capacity_hours`). |
| Total Availability | Net time after all leaves (`availability_hours` sum). |
| Total Planned Leaves | Sum of `planned_taken_hours + planned_not_taken_hours` across all members. |
| Total Unplanned Leaves | Sum of `unplanned_taken_hours` across all members. |
| Total Planned Hours | Sum of `subtask_planned_hours` (or `planned_hours`) across all members — subtask-level estimates only. |
| Total Logged Hours | Sum of `logged_hours` across all members. |
| Total Epics Assigned | Count of unique epic keys across all member `epic_keys` arrays. |
| Utilization vs Availability | Three progress bars: Planned %, Logged %, Leave rate (all relative to total availability). |
| Per Member Breakdown | Scrollable table of every member with their individual Availability / Planned / Logged values. |

**Behaviour notes:**
- Drawer respects the current Stats unit (Days / Hours). Switching the toggle while the drawer is open updates values immediately.
- Reloading team data while the drawer is open auto-refreshes the drawer content.
- Close with `×` button, clicking the backdrop, or pressing **Escape**.

## Business Logic

- All aggregate stats are derived entirely from the already-loaded `S.teamData` (the `/api/team-capacity-planner/data` response). No additional API call is made when opening the drawer.
- Planned hours use `subtask_planned_hours` when available, falling back to `planned_hours`. This mirrors the per-member card logic.
- Epic count is deduplicated across all members — a single epic shared by three members counts as 1.
- Utilization percentages are capped at 100% for display.

## Business Cases

- **Sprint / monthly planning review** — at a glance see how much of the team's available time is already committed.
- **Leave impact awareness** — compare planned vs unplanned leave totals to flag attendance risk.
- **Capacity vs commitment gap** — compare Total Availability against Total Planned Hours to spot over-allocation or idle capacity.
- **Epic portfolio view** — see how many distinct epics the team is spread across.

## Script Files

| File | Role |
| --- | --- |
| `team_capacity_planner.html` | Canonical source — all UI, CSS, and JS in a single file. |
| `report_html/team_capacity_planner.html` | Promoted copy served by the report server; auto-synced from canonical via `_promote_report_html_if_newer`. |
| `report_server.py` | Serves `/settings/team-capacity-planner` and all `/api/team-capacity-planner/*` routes. Handles promotion logic. |

## Dependent & Impacted Files

| File | Relationship |
| --- | --- |
| `report_server.py` | Serves route and promotes canonical HTML on access. |
| `ASSIGNEE_HOURS_CAPACITY.md` | Root-level operational doc; references this screen for context on capacity calculations. |
| `docs/capacity-user-guide/00-capacity-overview.md` | Module overview; should list Team Capacity Planner as a sub-screen. |

## Table Schema

The drawer reads data entirely from the `/api/team-capacity-planner/data` response (in-memory `S.teamData`). Key fields consumed per member:

| Field | Source | Meaning |
| --- | --- | --- |
| `capacity_hours` | DB capacity profile × working days | Gross hours in range |
| `availability_hours` | capacity − all leave hours | Net usable hours |
| `planned_taken_hours` | Leave rows | Planned leave already taken |
| `planned_not_taken_hours` | Leave rows | Planned leave upcoming |
| `unplanned_taken_hours` | Leave rows | Unscheduled absence |
| `subtask_planned_hours` | Canonical DB subtask estimates | Sum of subtask original estimates |
| `planned_hours` | Fallback if `subtask_planned_hours` absent | Alternative planned load field |
| `logged_hours` | Canonical worklogs | Actual time logged in range |
| `epic_keys` | Canonical assignments | Epics linked to the member's subtasks |

## Data Flow

1. User selects team + date range → clicks **Load** → `loadTeamData()` calls `/api/team-capacity-planner/data`.
2. Response is stored in `S.teamData`; `renderMembers()` paints the left panel.
3. User clicks **Analytics** button → `openAnalyticsDrawer()` → `renderAnalyticsStats()` aggregates totals from `S.teamData.members` in-memory — no extra fetch.
4. `renderAnalyticsStats()` writes stat cards, utilization bars, and per-member table into `#analytics-drawer-body`.
5. Changing the Days/Hours toggle calls `renderMembers()` and, if drawer is open, also `renderAnalyticsStats()` to reformat all values.
