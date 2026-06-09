# Team Capacity Planner

## Screen

- Name: Team Capacity Planner
- Route: `/settings/team-capacity-planner`
- Purpose: Review team capacity, leave impact, logged work, assigned epics, and per-resource planned subtask load.

## Sections

### Toolbar

The top bar is compacted into two viewport-aware dropdown menus (**Filters** and **Settings**) plus **Analytics** and **Load** buttons. Each dropdown is clamped inside the browser viewport so labels, toggles, and selectors remain visible at narrow widths.

#### Filters menu

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Team (`team-select`) | Dropdown | Yes | First configured performance team | Chooses the resource group to display. | Values come from `performance_teams`. |
| Date Range (dual calendar) | Custom date-range picker | Yes | First day → last day of current month | Single button showing `DD Mon YYYY → DD Mon YYYY`. Click opens a two-month calendar panel. Click a start day, then an end day — range auto-applies and data reloads. Quick-select shortcuts: **This month**, **Last month**, **This quarter**. | Hidden `from-date` / `to-date` inputs are updated on apply; JS reads them via `$fromDate.value` / `$toDate.value`. |

#### Settings menu

| Field | Type | Required | Default | Description | Validation / Rules |
| Capacity Profile (`profile-select`) | Dropdown | No | Auto best match | Selects the working-day profile used to convert hours to capacity/day stats. Now inside the **Settings** menu. | Auto mode picks the best matching saved profile. |
| Stats (`stat-unit-toggle`) | Animated pill toggle | No | Days | Switches resource stats between days and hours. A sliding white pill animates between the **Days** and **Hours** labels to show the active unit. Selection is persisted to `localStorage` key `tcp-stat-unit`. | Toggle is display-only; it does not reload or change saved data. |
| Epics by (`epic-mode-toggle`) | Animated pill toggle | No | **TK Dates** | Chooses which dates decide which epics are eligible for Planned / Logged calculations and for each member's Assigned Epics list. **TK Dates** (default) = an epic qualifies when at least one of its **stories** matches the Date Match rule. **Subtask Dates** = an epic qualifies when at least one of its **subtasks** matches the Date Match rule. With default Date Match = Overlap, date ranges only need to overlap the selected From-To window. | Selection is persisted to `localStorage` key `tcp-epic-mode` (`story` or `subtask`). Toggling auto-reloads team data and any open Subtask Breakdown drawer. |
| Date Match (`date-match-toggle`) | 3-option animated pill toggle | No | **Overlap** | Controls how epic/story/subtask dates are compared to the filter range. **Overlap** (default) = include if the item's date range overlaps the filter at all (`start_date ≤ filter_end AND due_date ≥ filter_start`). **Start in Range** = include only if the item's `start_date` falls within `[filter_start, filter_end]`. **End in Range** = include only if the item's `due_date` falls within `[filter_start, filter_end]`. | Selection is persisted to `localStorage` key `tcp-date-match` (`overlap`, `start_in_range`, or `end_in_range`). Toggling auto-reloads team data and any open Subtask Breakdown drawer. |
| Exclude Bug Subtasks (`exclude-bugs-toggle`) | Checkbox / switch | No | **ON** (checked) | When ON, bug sub-tasks are excluded from Planned, Logged, and the Subtask Breakdown drawer. When OFF, bug sub-tasks are included in all three. | Persisted to `localStorage` key `tcp-exclude-bugs` (`'1'` = exclude, `'0'` = include). Toggling auto-reloads team data and any open Subtask Breakdown drawer. |
| Analytics (`analytics-btn`) | Icon + label button | No | — | Opens the Team Analytics right-side drawer showing aggregate team stats. | Amber-orange gradient button. Active only after team data is loaded. Respects current stat unit (Days/Hours). |
| Load (`load-btn`) | Button | Yes | — | Fetches team capacity data for the selected team and date range. | Disabled while loading. |

### Team Members

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Planned | Bar/value | No | `0d` | Planned load for the resource. | **Epic-scoped:** only counts subtasks that roll up (directly via `epic_key`, or indirectly via parent story -> `epic_key`) to an eligible epic. With default Date Match = Overlap, eligible story/subtask date ranges only need to overlap the selected From-To window. For those subtasks belonging to the assignee, the original estimates are summed. Whether bug subtasks count depends on the **Exclude Bug Subtasks** toggle (default ON -> bugs excluded). Assigned epic and story estimates are always ignored. Click the bar to open the Subtask Breakdown drawer. |
| Logged | Bar/value | No | `0d` | Logged work for the resource. | **Epic-scoped:** uses the same epic-in-range filter as Planned — only subtasks whose epic's `start_date` AND `due_date` both fall inside the selected From–To window are eligible. For those subtasks, canonical worklogs in the date range by this assignee are summed. Whether bug subtasks count depends on the **Exclude Bug Subtasks** toggle (default ON → bugs excluded). Click the bar to open the Subtask Breakdown drawer. |
| Available | Bar/value | No | Derived | Remaining capacity after leave impact. | Uses selected/auto capacity profile and leave rows. |
| Assigned Epics | Expandable list | No | Empty | Epics associated with the resource's assigned subtasks **after the epic-in-range filter is applied**. Only epics that (a) qualify under the current **Epics by** mode (TK Dates / Subtask Dates) **and** (b) contain at least one subtask assigned to this member (respecting the Exclude Bug Subtasks toggle) are listed. | Used for navigation/highlighting only; not used directly in the Planned value. |

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

### Subtask Breakdown Drawer

Opened by **clicking the Planned or Logged bar/row** in any member card. Slides in from the right (760px wide) over a dark backdrop. Lists the exact subtasks that contribute to the selected bar value, so users can audit and reconcile the numbers.

| Section | Description |
| --- | --- |
| Header | Shows `<Planned\|Logged> Hours · <Member Name>` and a subtitle with date range. |
| Totals strip | Subtask count, total planned hours, and total logged-in-range hours for the listed subtasks. |
| Subtasks table | One row per subtask. Columns: Expand arrow, Subtask (key + summary + parent Story / Epic chip), Start, Due, Estimate, Logged. Clicking the issue key opens Jira in a new tab. |
| Worklogs subtable | Expandable per row (click anywhere on the row to toggle). Shows date, author, and hours for every worklog this assignee logged against the subtask inside the date range. |

**Backend source:** `GET /api/team-capacity-planner/member-subtasks?assignee=<name>&from=<YYYY-MM-DD>&to=<YYYY-MM-DD>&kind=<planned|logged>&exclude_bugs=<1|0>&epic_mode=<story|subtask>`.

- `kind=planned`: returns assigned subtasks whose **epic's start AND due both fall inside the From–To window** (matches the Planned bar's SQL — epic-scoped, not subtask-date-scoped).
- `kind=logged`: returns subtasks the assignee logged time against in the range, **restricted to subtasks whose epic's start AND due both fall inside the From–To window** (matches the Logged bar's SQL).
- `exclude_bugs=1` (default) applies the `lower(issue_type) NOT LIKE '%bug%'` filter to the subtask set. Passing `0` removes that filter so bug sub-tasks are included.
- Both kinds attach the worklog list (worklogs in range, by this assignee) and parent Story / Epic summaries when available.
- The RLT project is always excluded (TCP card values, drawer rows, and worklog totals).

**Behaviour notes:**
- Close with `×` button, clicking the backdrop, or pressing **Escape**.
- Expand arrow indicates rows that have logged worklogs; rows with no in-range worklogs still appear (for planned breakdown) but are not expandable.

## Business Logic

### Epic-in-range rule (Planned + Logged + Assigned Epics)

Planned, Logged, and the per-member Assigned Epics list all apply the same **epic-in-range** filter before any aggregation. The filter depends on the **Epics by** toggle in the top bar:

#### Mode A — `Epics by TK Dates` (default)

An epic qualifies when **at least one of its stories** matches the active Date Match rule. With default Date Match = Overlap, the story date range qualifies when `start_date <= filter_end AND due_date >= filter_start`.

SQL (essence):
```sql
SELECT DISTINCT upper(epic_key) FROM canonical_issues
WHERE lower(issue_type) LIKE '%story%'
  AND upper(project_key) != 'RLT'
  AND epic_key != ''
  AND start_date BETWEEN :from AND :to
  AND due_date   BETWEEN :from AND :to
```

#### Mode B — `Epics by Subtask Dates`

An epic qualifies when **at least one of its subtasks** matches the active Date Match rule. With default Date Match = Overlap, the subtask date range qualifies when `start_date <= filter_end AND due_date >= filter_start`. Subtask -> epic linkage is resolved either directly via `subtask.epic_key` or indirectly via `subtask.parent_issue_key` -> parent story's `epic_key`.

#### Common downstream behaviour (both modes)

1. Let `EligibleEpics` = the set of epic keys returned by the SQL above for the active mode.
2. Subtasks roll up to an eligible epic if either:
   - the subtask's own `epic_key` ∈ `EligibleEpics`, **or**
   - the subtask's parent story's `epic_key` ∈ `EligibleEpics`.
3. **Planned hours (per member)** = sum of `original_estimate_hours` of those subtasks where `lower(assignee) = member` (bug filter optional via toggle).
4. **Logged hours (per member)** = sum of `canonical_worklogs.hours_logged` of those subtasks where `issue_assignee = member` AND `started_date` ∈ [From, To] (bug filter optional).
5. **Assigned Epics (per member)** = subset of `EligibleEpics` where at least one matching subtask is assigned to the member (respects the Exclude Bug Subtasks toggle).
6. If `EligibleEpics` is empty, every member's Planned / Logged is `0` and Assigned Epics list is empty.

Example (TK Dates mode): User selects 1 Mar → 1 Dec. A story under epic `X` has start=1 Jun, due=1 Jul → epic `X` qualifies. A story under epic `Y` has start=1 Feb, due=1 May → epic `Y` is excluded because story start falls before Mar.

### Analytics drawer aggregates

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
| `team_capacity_planner.html` | Canonical source — all UI, CSS, and JS in a single file. Hosts both the Analytics drawer and the new Subtask Breakdown drawer. |
| `report_html/team_capacity_planner.html` | Promoted copy served by the report server; auto-synced from canonical via `_promote_report_html_if_newer`. |
| `report_server.py` | Serves `/settings/team-capacity-planner` and all `/api/team-capacity-planner/*` routes including `/member-subtasks`. Implements the per-member Planned (`_tcp_build_team_data` → `subtask_planned_hours` SQL) and Logged (joined to `canonical_issues` to exclude bugs) calculations and the subtask breakdown endpoint. |

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

## Change Notes

- 2026-06-02: Team data aggregation now batches planned-hours, logged-hours, and assigned-epic lookups across all selected team members instead of running the same canonical queries once per member.
- 2026-06-02: Planner API calls use a shared JSON fetch helper so Azure gateway timeout or application-error pages show an HTTP/non-JSON diagnostic instead of a browser JSON parse error.
