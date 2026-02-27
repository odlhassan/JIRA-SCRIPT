# Assignee Hours Capacity Integration

## Screen

- Name: Assignee Hours Report - Capacity Planning section
- Route: `/assignee_hours_report.html`
- Purpose: Calculate capacity, save/load reusable profiles, and display leave-adjusted capacity KPIs.

## Sections

### Capacity Planning

#### Area: Input Form

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| From Date | Date | Yes | Current report filter start | Calculation range start. | Must be valid and not after To Date. |
| To Date | Date | Yes | Current report filter end | Calculation range end. | Must be valid and not before From Date. |
| Employees | Number | Yes | `0` | Team size for capacity. | Must be `>= 0`. |
| Standard Hours/Day | Number | Yes | `8` | Hours for non-Ramadan weekdays. | Must be `> 0`. |
| Ramadan Start | Date | Conditional | Empty | Ramadan period start. | Must be paired with Ramadan End. |
| Ramadan End | Date | Conditional | Empty | Ramadan period end. | Must be paired with Ramadan Start and not earlier than start. |
| Ramadan Hours/Day | Number | Yes | `6.5` | Hours for Ramadan weekdays. | Must be `> 0`. |
| Holiday Dates | Multi-date list | No | Empty | Excluded dates. | ISO date format; duplicates removed. |

#### Area: Profile Reuse

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Reuse Saved Capacity | Dropdown | No | Empty/first profile | Shows saved ranges from `/api/capacity/profiles`. | Disabled/limited in static mode (`file://`). |
| Apply Profile To Current Range | Button | No | Enabled | Loads selected profile settings into current range context. | Requires valid selected profile. |
| Save Capacity | Button | No | Enabled | Saves current settings through `POST /api/capacity`. | Upsert by `from_date + to_date`. |
| Recalculate | Button | No | Enabled | Recomputes metrics using server calculate API or client fallback. | Uses current form values. |

### KPI Cards

#### Area: Capacity and Leave Outputs

##### Fields

| KPI | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Total Capacity | Calculated numeric | No | `0h` | Profile-based available capacity. | Computed using weekday, holiday, and Ramadan logic. |
| Leave Hours | Calculated numeric | No | `0h` | Sum of taken and not-yet-taken leave hours. | Derived from leave report data. |
| Remaining Capacity | Calculated numeric | No | `0h` | `Total Capacity - Leave Hours`. | Updates after recalc/apply/save. |
| Capacity After Leaves | Calculated numeric | No | `0h` | `Available Capacity - Project Actual Hours - Leave Hours`. | Excludes `RLT` project from project actual hours. |

## Logic and API Linkage

- Capacity profile list: `GET /api/capacity/profiles`
- Save profile: `POST /api/capacity`
- Load profile by range: `GET /api/capacity?from=...&to=...`
- Recalculate: `POST /api/capacity/calculate` (with client-side fallback in static mode)
- Data persistence target: `assignee_hours_capacity.db`

## Integration Notes

- This report is the primary editor for calculation output behavior.
- Saved profiles created here are directly reusable on:
  - Capacity Settings page
  - Nested View report profile selector

