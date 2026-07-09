# RnD Muscle Utilization Settings

## Purpose

`RnD Muscle Utilization` is an admin planning page for managers who map RnD resources to Epics Planner work. The configuration route shows the epic catalog, the central planner, and the resource catalog. The view route shows only the central planner for review mode.

## Current Scope

The page loads live state through `/api/rnd-muscle-utilization`, supports skill/team/project configuration, and persists planner epics, resource mappings, and card order to SQLite tables in the dedicated RnD Muscle SQLite database. Planner moves and resource mappings update the browser immediately, then auto-save after 2.5 seconds without another planner action. Those tables are auto-initialised on first call.

## Business Logic

- **Skills** - The 11 default skills from `DEFAULT_RND_MUSCLE_SKILLS` are seeded once on first DB initialisation. Managers can add custom skills. Duplicate names are rejected case-insensitively.
- **Teams** - Managers create color-coded teams, assign skills and resources to them. Team skills describe the collective capability of the team, not mandatory skills for every team member. Team color must be selected from the fixed 10-color UI-friendly palette; duplicate team names are rejected.
- **Resources** - Resources are stored with display name, initials, email, optional team membership, and explicit resource skills. Canonical assignee/worklog names are synced into `rnd_muscle_resources` on page-state load. The right resource catalog hides resources marked resigned in the source capacity database, groups the remaining active resources by assigned team, and keeps resources without a team under **No team**. The catalog search filters active resources while the user types across name, initials, email, and team name. Resource catalog rows use a light tone derived from the assigned team's color, with the original color preserved as an accent; resources without a team keep the neutral row treatment. Resource skill counts include only the skills selected directly for that resource.
- **Epics** - The epic catalog is sourced from `epics_management`. Priority comes from the `priority` column. Budgeted hours are derived from `epics_management_plan_values.epic_plan` as `man_days * 8`.
- **Mapped epics** - The left planner area is the master collection list, similar to Firestore collections. It holds mapped/planned epic cards. Epics enter by dragging from the Epics catalog, clicking **Planner** in the catalog row, or clicking **Add to planner** from the backlog table. While a mapped epic card is being dragged for sorting, a dashed empty placeholder appears between cards or at the end of the list to show the exact drop position. The browser applies the new order instantly and queues the latest `sort_order` payload for `/api/rnd-muscle-utilization/planner/reorder` after the 2.5-second idle window.
- **Mapped resources** - The right planner area is the detail document list for the selected mapped epic, similar to Firestore documents. Dropping a resource from the right catalog onto this area maps it to the selected epic immediately in the UI. Resource cards are colored with their team's configured color and can be dragged to reorder per epic. Mapping, unmapping, and resource-card reorder actions replace the pending save payload for that epic and persist only after the user pauses for 2.5 seconds. When an active resource is already mapped to one or more planner epics, the right catalog row shows a **booked on n epics** chip; clicking it opens a small menu listing the booked epic keys, names, projects, and allocation hours.
- **Save progress** - A compact progress bar appears under the status line whenever planner changes are pending or actively saving. It advances while waiting for the idle window, continues while the database request is running, turns green on success, and turns red if auto-save fails. Text-only save feedback is not the only indicator.
- **Backlog** - The backlog table sits under the mapping section and is separate from mapped planner epics. It shows epics that should be planned next and provides **Remove** and **Add to planner** actions for each row. Users can add to backlog by clicking **Backlog** in the left catalog, dragging an epic from the left catalog into the backlog table, or dragging an already mapped epic card into the backlog table.
- **Mappings** - Epic-to-resource assignments are upserted by `(epic_key, resource_id)` in `rnd_muscle_epic_resource_mappings`. The table stores `allocation_hours` and `sort_order`. Removing a resource card unmaps that resource from the selected epic.
- **Planner views** - **Hierarchical** shows exactly two planner areas: compact stacked mapped-epic cards and mapped-resource cards for the selected epic. Mapped-epic sorting keeps the dashed insertion placeholder stable while the pointer passes over card gaps or the placeholder itself. The Cluster panel is hidden until the user explicitly chooses **Cluster**. **Cluster** shows mapped epics on the left, resource-initial bubbles huddled together by default, team-color legend entries, hover details with full resource name and team, and animated bubble alignment with stronger node connections when a user selects an epic.
- **View Mode route** - The **View Mode** button opens `/settings/rnd-muscle-utilization/view`. That page uses the same planner state and APIs but hides side epic/resource catalogs and configuration controls.
- **Theme controls** - The shared toolbar has a two-state **Dark / Light** toggle switch and **Color** picker. Choices are stored in browser local storage under `rnd-muscle-theme-mode-v1` and `rnd-muscle-theme-color-v1`, so the configuration and view routes use the same VS Code-style chrome, spacing, compact borders, accent color, and themed scrollbars.
- **Quick stats** - Computed on every page-state load: resources assigned to any epic, resources not yet assigned, total epics in the selected project scope, and high-priority epics with no resource assignment.
- **Project tabs** - The ALL tab appears first. The Configure Projects modal uses project checkboxes; the visible tab strip is limited to that selection.

## Screen Areas

| Area | Fields And Controls | Behavior |
| --- | --- | --- |
| Quick stats band | Associated resources, unassociated resources, high-priority unassigned epics, active project epic count | Computed live from DB on every load. Hidden on `/settings/rnd-muscle-utilization/view`. |
| Left panel | Epic search, project filter, scrollable epic catalog | Filtered by `search_rnd_muscle_epics`. Each epic row has a **Planner** button and is draggable into the planner epic area. Hidden on the view route. |
| Middle panel | Project tabs, Hierarchical/Cluster toggle, mapped epic area, mapped resource area, save progress bar, cluster canvas, backlog table | Hierarchical mode has only mapped epics and mapped resources in the mapping section. Epic card sorting shows a dashed insertion placeholder and updates local order immediately; hovering through that placeholder preserves the current insertion position so the list does not reflow unexpectedly. Planner order and resource mappings are auto-saved after a 2.5-second idle delay, with the progress bar visible while pending or saving. The backlog table appears below the mapping section as the plan-next queue and accepts dragged epics from the left catalog or mapped epic cards. Cluster mode switches to the epic list plus animated resource bubble canvas. |
| Right panel | Resource search, team-grouped active resource list, team/skill badges, booked-epic chip, Edit skills action | Resources are loaded from `rnd_muscle_resources`; names found in `performance_resource_resignations` are filtered out of the catalog. The remaining resources are grouped by assigned team, with unassigned people under **No team**. Search filters while typing across resource name, initials, email, and team. Assigned-team rows use the team color, while no-team rows stay neutral. The skill badge counts only skills selected directly for that resource. A **booked on n epics** chip appears when the resource has existing `rnd_muscle_epic_resource_mappings`; clicking the chip shows the booked epics. Dragging a resource into the selected-epic resource area maps it. Hidden on the view route. |
| Theme controls | Dark/Light switch, accent color picker | Applies a VS Code-style light/dark theme and accent color to both configuration and view mode through shared local storage keys. Scrollbar tracks follow the active panel theme and scrollbar thumbs follow the selected accent color. |
| Configuration controls | View Mode, Create Team, Add Skill, Configure Projects | Backed by the RnD APIs and local project-tab selection. The view route replaces these with a **Configuration** link while keeping theme controls visible. |

## Script Files

| File | Role |
| --- | --- |
| `rnd_muscle_utilization_types.py` | Strict dataclasses, typed payloads, default skills, and page-state shape. |
| `rnd_muscle_utilization_service.py` | Dedicated RnD DB init, legacy table migration helper, page-state load, canonical resource sync from the source DB, search, team/skill CRUD, planner epic management, mapping persistence, and reorder persistence. |
| `report_server.py` | Registers `/settings/rnd-muscle-utilization`, `/settings/rnd-muscle-utilization/view`, RnD API routes, and the interactive page HTML. |

## Table Schema

| Table | Key Columns | Purpose |
| --- | --- | --- |
| `rnd_muscle_skills` | skill_id PK, name, is_default | Default and manager-defined custom skills. |
| `rnd_muscle_teams` | team_id PK, name, color_hex | Manager-defined color-coded skillset teams. |
| `rnd_muscle_team_skills` | (team_id, skill_id) PK | Many-to-many team-to-skill associations. |
| `rnd_muscle_resources` | resource_id PK, display_name, initials, email, team_id | People available for epic assignment. |
| `rnd_muscle_resource_skills` | (resource_id, skill_id) PK | Many-to-many resource-to-skill associations. |
| `rnd_muscle_planner_epics` | epic_key PK, sort_order | Epics explicitly added to the planner epic area, ordered by saved card order. |
| `rnd_muscle_backlog` | epic_key PK, sort_order | Epics queued for later planning, rendered under the mapping section with Remove and Add to planner actions. |
| `rnd_muscle_epic_resource_mappings` | (epic_key, resource_id) PK, allocation_hours, sort_order | Resource-to-epic assignment with optional hour allocation and selected-epic resource card order. |

By default, `rnd_muscle_*` tables stay in the existing capacity DB (`assignee_hours_capacity.db`) so existing production team/resource configuration remains available. If `JIRA_RND_MUSCLE_UTILIZATION_DB_PATH` is explicitly configured, those tables are stored in that dedicated feature DB (`rnd_muscle_utilization.db` is the conventional filename) and auto-created by `_init_rnd_muscle_utilization_db`.

The feature reads Epics Planner and canonical Jira data from the source capacity DB (`assignee_hours_capacity.db` by default, or `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH` when configured). When a separate RnD DB is explicitly configured, `report_server.py` passes both paths to the service: the RnD DB for feature-owned writes and the capacity DB for read-only source lookups.

## Data Flow

1. **Page load** - `load_rnd_muscle_utilization_page_state(rnd_db_path, source_db_path)` initialises feature schema, syncs canonical resource names from the source DB, reads RnD tables plus source `epics_management`, `epics_management_plan_values`, and `performance_resource_resignations`, and returns `RndMuscleUtilizationPageState`.
2. **Epic search** - `search_rnd_muscle_epics(rnd_db_path, text, project_keys, source_db_path)` filters the epic catalog by substring and project key.
3. **Team CRUD** - `save_rnd_muscle_team` validates that color is one of the 10 supported palette values and rejects duplicate names, upserts the team, replaces team skills, updates resource `team_id`, and returns fresh state.
4. **Skill CRUD** - `add_rnd_muscle_skill` rejects duplicates and inserts custom skills with `is_default=0`.
5. **Resource skill mapping** - `save_rnd_muscle_resource_skills` validates resource and skill IDs, replaces direct rows in `rnd_muscle_resource_skills`, and returns fresh state. The browser lets managers choose from the full skill list for each resource, independent of team skills.
6. **Backlog** - `add_epic_to_rnd_muscle_backlog` queues epics for future planning. `reorder_rnd_muscle_backlog` saves backlog order. `remove_epic_from_rnd_muscle_backlog` removes only the backlog row and does not alter active planner mappings.
7. **Planner epics** - `add_epic_to_rnd_muscle_planner` adds active mapped epics. The browser adds planner cards locally before the API call completes, then queues the POST for the idle-save cycle. `reorder_rnd_muscle_planner_epics` saves drag order. `remove_epic_from_rnd_muscle_planner` deletes the planner row and dependent resource mappings.
8. **Mapping** - `save_rnd_muscle_epic_resource_mapping` removes no-longer-assigned resources, upserts remaining assignments with allocation hours and sequential `sort_order`, and returns fresh state. Mapping, unmapping, and selected-epic resource ordering now share a debounced browser save queue keyed by epic, so rapid consecutive changes collapse into the latest payload before SQLite is updated.
9. **Mapped epic sorting** - Dragging a mapped epic shows a dashed insertion placeholder before or after the hovered mapped-epic card. Cards and placeholders use short transitions for a smoother drag feel. When the cursor crosses the placeholder or other non-card gaps inside the mapped-epics list, the browser keeps the existing placeholder target until the user reaches another epic card or the drop zone end. Dropping the card reorders local state immediately and queues `/api/rnd-muscle-utilization/planner/reorder`.
10. **Tab strip** - Server state exposes all project tabs; the browser stores selected tab projects in local storage and renders only the selected tabs.
11. **Theme preference** - The browser reads and writes `rnd-muscle-theme-mode-v1` and `rnd-muscle-theme-color-v1`, then updates CSS variables on load and whenever a user changes the toolbar controls. Scrollable page areas use the same CSS variables so vertical scrollbars change with the selected theme and accent color.

## Schema Change Notes

- Added `rnd_muscle_epic_resource_mappings.sort_order INTEGER NOT NULL DEFAULT 0` to preserve resource card order per selected epic.
- Added `rnd_muscle_planner_epics` to separate active mapped epics from the backlog plan-next queue.
- Moved all feature-owned `rnd_muscle_*` tables out of `assignee_hours_capacity.db` and into `rnd_muscle_utilization.db`. The capacity DB remains the read source for Epics Planner and canonical Jira data.
- `migrate_legacy_rnd_muscle_tables(legacy_db_path, rnd_db_path, drop_legacy=True)` copies existing local feature data into the new DB and removes legacy tables after row-count verification.
- The local changelog and schema snapshot were updated. Production migration is pending until a human-downloaded production DB file is provided locally.

## Related Import Change

Epics Planner Import expects and seeds the three support effort phases as planner columns: `ReadAPI Sup`, `SiteLayout Sup`, and `OmniAgent Sup`. These phases are treated like other most-likely input columns and can carry zero or positive man-days for every imported epic.
