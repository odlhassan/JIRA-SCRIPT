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

### Team Members

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Planned | Bar/value | No | `0d` | Planned load for the resource. | Sums only assigned subtask original estimates in the selected date range. Assigned epic and story estimates are intentionally ignored. |
| Logged | Bar/value | No | `0d` | Logged work for the resource in the selected date range. | Uses canonical worklogs by issue assignee. |
| Available | Bar/value | No | Derived | Remaining capacity after leave impact. | Uses selected/auto capacity profile and leave rows. |
| Assigned Epics | Expandable list | No | Empty | Epics associated with the resource's assigned work items. | Used for navigation/highlighting only; not used directly in the Planned value. |
