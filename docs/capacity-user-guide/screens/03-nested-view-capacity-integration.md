# Nested View Capacity Integration

## Screen

- Name: Nested View Report - Capacity Profile section
- Route: `/nested_view_report.html`
- Purpose: Apply an existing saved capacity profile to report KPIs without editing profile definitions.

## Sections

### Capacity Profile

#### Area: Profile Controls

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Saved Capacity Profiles (`capacity-profile-select`) | Dropdown | No | First saved profile is auto-selected when available | Shows all saved profiles by range and parameters. | Disabled when none exist. |
| Apply (`capacity-profile-apply`) | Button | No | Enabled when profiles exist | Applies selected profile to current report date filter. | Requires selected profile. |
| Refresh (`capacity-profile-refresh`) | Button | No | Enabled | Reloads profiles from `GET /api/capacity/profiles`. | Server mode required for latest persisted data. |
| Use Project Totals (`capacity-profile-reset`) | Button | No | Enabled | Clears applied profile override. | Reverts KPIs to project-total capacity logic. |
| Manage Capacity Profiles | Link button | No | Enabled | Opens settings page for create/edit/delete. | Navigates to `/settings/capacity`. |
| Status (`capacity-profile-status`) | Read-only text | No | Informational message | Shows load/apply/reset errors and confirmation. | Variant state: info/success/error. |
| Profile Details (`capacity-profile-details`) | Read-only text | No | Guidance text | Shows selected profile details and computed capacity. | Includes dynamic capacity when profile is applied. |

## KPI Linkage

- Performance Overview score card order:
  - `Total Capacity (Hours)`
  - `Total Planned Projects (Hours)`
  - `Total Leaves Planned`
  - `Total Actual Project Hours`
  - `Availability`
  - `Hours Required To Complete Projects`
  - `Capacity Gap`
- `Total Capacity (Hours)` card behavior:
  - Formula: `Total Capacity (Hours) = Employee Count x Available Business Days x Per Day Hours`.
  - Source:
    - On load, first available saved profile is auto-applied and used against current filter bounds.
    - If no profile exists (or after `Use Project Totals`), falls back to project-total capacity.
  - Top-right formula chip shows icon-based live values:
    - `person x calendar_month x hourglass_top`
    - Employee count, business weekdays, and computed per-day hours.
  - `i` tooltip displays formula and selected-profile ingredient values (employees, per-day hours, weekday split, holidays, and calculated profile capacity).
- `Total Planned Projects (Hours)` card behavior:
  - Formula: `Sum(Epic Original Estimate Hours), where Epic Planned Start OR Planned End is in selected date range, excluding RLT (RnD Leave Tracker)`.
- `Total Actual Project Hours` card behavior:
  - Formula: `Sum(Project Actual Hours), excluding RLT (RnD Leave Tracker)`.
- `Hours Required To Complete Projects` card behavior:
  - Formula: `Total Planned Projects - Total Actual Project Hours`.
- `Total Leaves Planned` card behavior:
  - Formula: `Sum(Original Estimates) for RLT leave work`.
- `Availability` card behavior:
  - Formula: `Total Capacity (Hours) - Total Leaves Planned`.
- `Capacity Gap` card behavior:
  - Formula: `Total Capacity (Hours) - Total Planned Projects (Hours) - RLT RnD Leave Tracker Original Estimates`.
  - Recalculates when date filters, project filters, or capacity profile selection changes.
- Performance Overview `i` icon behavior:
  - Shows formula text and live ingredient values for each KPI card.
  - Supported visible cards: `Total Capacity`, `Total Planned Projects`, `Total Actual Project Hours`, `Hours Required To Complete Projects`, `Total Leaves Planned`, `Availability`, and `Capacity Gap`.
  - `Total Leaves Taken` card is currently commented out in UI (kept in code for later use).

## Logic and API Linkage

- Profile source on load:
  - Embedded `capacity_profiles` in report payload.
  - First available profile is auto-applied when present.
- Live refresh source:
  - `GET /api/capacity/profiles`
- Settings handoff:
  - `/settings/capacity` for managing profile definitions.

## Constraints

- Nested View does not save/edit/delete profiles directly.
- It only applies or clears a profile override for visualization/calculation on this report page.
