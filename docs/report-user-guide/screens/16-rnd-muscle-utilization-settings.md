# RnD Muscle Utilization Settings

## Purpose

`RnD Muscle Utilization` is an admin planning page for managers who map RnD resources to Epics Planner work. The configuration route shows the epic catalog, the central planner, and the resource catalog. The view route shows only the central planner for review mode.

## Current Scope

The page loads live state through `/api/rnd-muscle-utilization`, supports skill/team/project configuration, and persists planner epics, resource mappings, and card order to SQLite tables that are auto-initialised on first call.

## Business Logic

- **Skills** - The 11 default skills from `DEFAULT_RND_MUSCLE_SKILLS` are seeded once on first DB initialisation. Managers can add custom skills. Duplicate names are rejected case-insensitively.
- **Teams** - Managers create color-coded teams, assign skills and resources to them. Team color must be a valid 3- or 6-digit hex code. Duplicate team names are rejected.
- **Resources** - Resources are stored with display name, initials, email, and optional team membership. Canonical assignee/worklog names are synced into `rnd_muscle_resources` on page-state load.
- **Epics** - The epic catalog is sourced from `epics_management`. Priority comes from the `priority` column. Budgeted hours are derived from `epics_management_plan_values.epic_plan` as `man_days * 8`.
- **Mapped epics** - The left planner area is the master collection list, similar to Firestore collections. It holds mapped/planned epic cards. Epics enter by dragging from the Epics catalog, clicking **Planner** in the catalog row, or clicking **Add to planner** from the backlog table. While a mapped epic card is being dragged for sorting, a dashed empty placeholder appears between cards or at the end of the list to show the exact drop position. The saved list is backed by `rnd_muscle_planner_epics`; `sort_order` is updated through `/api/rnd-muscle-utilization/planner/reorder`.
- **Mapped resources** - The right planner area is the detail document list for the selected mapped epic, similar to Firestore documents. Dropping a resource from the right catalog onto this area maps it to the selected epic. Resource cards are colored with their team's configured color and can be dragged to save per-epic order.
- **Backlog** - The backlog table sits under the mapping section and is separate from mapped planner epics. It shows epics that should be planned next and provides **Remove** and **Add to planner** actions for each row. Users can add to backlog by clicking **Backlog** in the left catalog, dragging an epic from the left catalog into the backlog table, or dragging an already mapped epic card into the backlog table.
- **Mappings** - Epic-to-resource assignments are upserted by `(epic_key, resource_id)` in `rnd_muscle_epic_resource_mappings`. The table stores `allocation_hours` and `sort_order`. Removing a resource card unmaps that resource from the selected epic.
- **Planner views** - **Hierarchical** shows exactly two planner areas: compact stacked mapped-epic cards and mapped-resource cards for the selected epic. The Cluster panel is hidden until the user explicitly chooses **Cluster**. **Cluster** shows mapped epics on the left, resource-initial bubbles huddled together by default, team-color legend entries, and animated bubble alignment with stronger node connections when a user selects an epic.
- **View Mode route** - The **View Mode** button opens `/settings/rnd-muscle-utilization/view`. That page uses the same planner state and APIs but hides side epic/resource catalogs and configuration controls.
- **Theme controls** - The shared toolbar has a **Theme** selector (`Dark` / `Light`) and **Color** picker. Choices are stored in browser local storage under `rnd-muscle-theme-mode-v1` and `rnd-muscle-theme-color-v1`, so the configuration and view routes use the same VS Code-style chrome, spacing, compact borders, accent color, and themed scrollbars.
- **Quick stats** - Computed on every page-state load: resources assigned to any epic, resources not yet assigned, total epics in the selected project scope, and high-priority epics with no resource assignment.
- **Project tabs** - The ALL tab appears first. The Configure Projects modal uses project checkboxes; the visible tab strip is limited to that selection.

## Screen Areas

| Area | Fields And Controls | Behavior |
| --- | --- | --- |
| Quick stats band | Associated resources, unassociated resources, high-priority unassigned epics, active project epic count | Computed live from DB on every load. Hidden on `/settings/rnd-muscle-utilization/view`. |
| Left panel | Epic search, project filter, scrollable epic catalog | Filtered by `search_rnd_muscle_epics`. Each epic row has a **Planner** button and is draggable into the planner epic area. Hidden on the view route. |
| Middle panel | Project tabs, Hierarchical/Cluster toggle, mapped epic area, mapped resource area, cluster canvas, backlog table | Hierarchical mode has only mapped epics and mapped resources in the mapping section. Epic card sorting shows a dashed insertion placeholder before saving the order; resource card order is also saved. The backlog table appears below the mapping section as the plan-next queue and accepts dragged epics from the left catalog or mapped epic cards. Cluster mode switches to the epic list plus animated resource bubble canvas. |
| Right panel | Scrollable resource list, team/skill badges | Resources are loaded from `rnd_muscle_resources`; dragging a resource into the selected-epic resource area maps it. Hidden on the view route. |
| Theme controls | Theme selector, accent color picker | Applies a VS Code-style light/dark theme and accent color to both configuration and view mode through shared local storage keys. Scrollbar tracks follow the active panel theme and scrollbar thumbs follow the selected accent color. |
| Configuration controls | View Mode, Create Team, Add Skill, Configure Projects | Backed by the RnD APIs and local project-tab selection. The view route replaces these with a **Configuration** link while keeping theme controls visible. |

## Script Files

| File | Role |
| --- | --- |
| `rnd_muscle_utilization_types.py` | Strict dataclasses, typed payloads, default skills, and page-state shape. |
| `rnd_muscle_utilization_service.py` | DB init, page-state load, canonical resource sync, search, team/skill CRUD, planner epic management, mapping persistence, and reorder persistence. |
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

All tables are stored in the shared settings DB (`assignee_hours_capacity.db`) and auto-created by `_init_rnd_muscle_utilization_db`.

## Data Flow

1. **Page load** - `load_rnd_muscle_utilization_page_state(db_path)` initialises schema, syncs canonical resource names, reads RnD tables plus `epics_management` and `epics_management_plan_values`, and returns `RndMuscleUtilizationPageState`.
2. **Epic search** - `search_rnd_muscle_epics(db_path, text, project_keys)` filters the epic catalog by substring and project key.
3. **Team CRUD** - `save_rnd_muscle_team` validates color and duplicate names, upserts the team, replaces team skills, updates resource `team_id`, and returns fresh state.
4. **Skill CRUD** - `add_rnd_muscle_skill` rejects duplicates and inserts custom skills with `is_default=0`.
5. **Backlog** - `add_epic_to_rnd_muscle_backlog` queues epics for future planning. `reorder_rnd_muscle_backlog` saves backlog order. `remove_epic_from_rnd_muscle_backlog` removes only the backlog row and does not alter active planner mappings.
6. **Planner epics** - `add_epic_to_rnd_muscle_planner` adds active mapped epics. `reorder_rnd_muscle_planner_epics` saves drag order. `remove_epic_from_rnd_muscle_planner` deletes the planner row and dependent resource mappings.
7. **Mapping** - `save_rnd_muscle_epic_resource_mapping` removes no-longer-assigned resources, upserts remaining assignments with allocation hours and sequential `sort_order`, and returns fresh state. `reorder_rnd_muscle_epic_resources` saves selected-epic resource card order.
8. **Tab strip** - Server state exposes all project tabs; the browser stores selected tab projects in local storage and renders only the selected tabs.
9. **Theme preference** - The browser reads and writes `rnd-muscle-theme-mode-v1` and `rnd-muscle-theme-color-v1`, then updates CSS variables on load and whenever a user changes the toolbar controls. Scrollable page areas use the same CSS variables so vertical scrollbars change with the selected theme and accent color.

## Schema Change Notes

- Added `rnd_muscle_epic_resource_mappings.sort_order INTEGER NOT NULL DEFAULT 0` to preserve resource card order per selected epic.
- Added `rnd_muscle_planner_epics` to separate active mapped epics from the backlog plan-next queue.
- The local changelog and schema snapshot were updated. Production migration is pending until a human-downloaded production DB file is provided locally.

## Related Import Change

Epics Planner Import expects and seeds the three support effort phases as planner columns: `ReadAPI Sup`, `SiteLayout Sup`, and `OmniAgent Sup`. These phases are treated like other most-likely input columns and can carry zero or positive man-days for every imported epic.
