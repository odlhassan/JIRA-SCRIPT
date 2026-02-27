# RnD Data Story

## Screen

- Name: RnD Capacity and Workload Data Story
- Route: `/rnd_data_story.html`
- Purpose: Provide leadership-ready visuals for capacity, workload, pending demand, and epic status over a selected date range.

## Sections

### Story Navigation

#### Area: Page Flow Controls

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Previous Page (`story-prev`) | Button | No | Disabled on page 1 | Moves to previous story page. | Disabled at first page. |
| Next Page (`story-next`) | Button | No | Enabled on page 1 | Moves to next story page. | Disabled at last page. |
| Page Label (`story-page-label`) | Read-only text | No | `Page 1 of 5` | Shows current page position in story flow. | Updates on every navigation. |

### Story Pages

- Page 1: Leadership Snapshot (all six KPIs)
- Page 2: Capacity vs Current Commitment (funnel: Available, After Leaves, Booked, Remaining Capacity, Hours Required to Complete Epics)
- Page 3: Pending Demand and Coverage (comparison bars + coverage ratio ring + estimate/logged/required/gap breakdown + top pending chart)
- Page 4: Execution Pressure (status split + narrative insight cards)
- Page 5: Executive Summary (decision-oriented text summary)

### Controls

#### Area: Date and Profile Filters

##### Fields

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| From Date (`from-date`) | Date | Yes | Payload default start date | Start date for KPI and chart filtering. | Must be valid and not after To Date. |
| To Date (`to-date`) | Date | Yes | Payload default end date | End date for KPI and chart filtering. | Must be valid and not before From Date. |
| Capacity Profile (`capacity-profile-select`) | Dropdown | No | First saved profile when available | Selects saved capacity settings for story calculations. | Disabled if no profiles exist. |
| Apply (`apply-btn`) | Button | No | Enabled | Recomputes KPIs/charts for selected date range and active profile. | Requires valid date range. |
| Reset (`reset-btn`) | Button | No | Enabled | Resets dates to defaults and clears profile override. | Restores default story state. |
| Profile Status (`profile-status`) | Read-only text | No | Informational text | Shows active profile feedback and apply/reset messages. | Informational only. |

### KPI Cards

#### Area: Leadership KPIs

##### Fields

| KPI | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Leave-Adjusted Capacity | Calculated numeric | No | `0h` | Capacity after applying leave deductions. | `available_capacity_hours - (planned_taken + planned_not_taken + unplanned_taken)` |
| Work On Plate | Calculated numeric | No | `0h` | Planned committed epic hours in selected range. | Sum of nested-view epic planned hours (`rmi.man_hours` fallback `rmi.man_days*8`) where epic `planned_start` OR `planned_end` is in range, excluding `RLT`. |
| Investable More Hours | Calculated numeric | No | `0h` | Additional hours available after workload. | `leave_adjusted_capacity - work_on_plate` |
| Pending Hours Required | Calculated numeric | No | `0h` | Remaining epic demand in range. | Sum of `max(epic_original_estimate_hours - epic_logged_hours, 0)` |
| Closed (Resolved) Epics | Calculated count | No | `0` | Epic count in resolved status family. | Status contains `resolved` (case-insensitive). |
| Open (In Progress) Epics | Calculated count | No | `0` | Epic count in active in-progress family. | Status contains `in progress` (case-insensitive). |

### Story Visuals

#### Area: Chart Cards

##### Fields

| Chart | Type | Description | Validation / Rules |
| --- | --- | --- | --- |
| Capacity Commitment Funnel | Funnel | Shows Available Capacity, After Leaves, Booked Work, Remaining Capacity, and Hours Required to Complete Epics in one comparable visual. | Uses same date filter and planned-hours commitment logic as KPI calculations. |
| Pending Required vs Investable Breakdown | Composite chart | Compares investable vs pending, shows coverage ratio, and breaks out Total Estimate, Actual Logged, Hours Required, and Coverage Gap. | Pending uses epic-level estimate minus logged clamp-to-zero; hours required uses total estimate minus total logged. |
| Epic Status Split | Split bar | Shows Resolved vs In Progress counts. | Counts only epics passing epic date filter rule. |
| Top Pending Epic Demand | Horizontal bar chart | Top epics by pending hours. | Sorted descending by pending hours; top N rendered. |

### Insight Callouts

#### Area: Narrative Output

##### Fields

| Insight | Type | Description | Validation / Rules |
| --- | --- | --- | --- |
| Capacity surplus/deficit | Text callout | Indicates positive/negative investable hours. | Uses `investable_more_hours` sign. |
| Pending demand coverage gap | Text callout | Indicates whether pending demand can be covered. | Uses `investable_more_hours - pending_required_hours`. |
| Epic status pressure | Text callout | Indicates pressure from open vs resolved counts. | Compares in-progress count against resolved count. |

## Logic and Filters

- Department scope is fixed: `Research and Development (RnD)`, covering all projects and all assignees.
- Epic inclusion for epic-based metrics and counts:
  - Include epic when `start_date` is in selected range OR `end_date` is in selected range.
  - Epics with both dates missing are excluded from epic-based metrics.
- Pending-hours formula:
  - `sum(max(epic_original_estimate_hours - epic_logged_hours, 0))`
- Hours-required-to-complete formula (page 2 funnel):
  - `sum(epic_original_estimate_hours) - sum(epic_logged_hours)` for filtered epics.
- Work-on-plate formula:
  - `sum(rmi.man_hours)` (fallback `rmi.man_days * 8`) for epics where `planned_start` OR `planned_end` is in selected range, excluding `RLT`.
  - This is intentionally aligned to Nested View `Total Planned Projects (Hours)` to keep both fields identical.

## Data Sources

- `1_jira_work_items_export.xlsx`: epic metadata, status, dates, original estimates.
- `2_jira_subtask_worklogs.xlsx`: epic logged-hours rollup via `parent_epic_id`.
- `assignee_hours_report.xlsx`: workload logs by date/assignee/project.
- `rlt_leave_report.xlsx`: leave metrics used in leave-adjusted capacity.
- `nested view.xlsx`: normalized epic planned hours used for `Work On Plate` parity with Nested View report.
- Capacity API/profile data:
  - `GET /api/capacity?from=...&to=...`
  - `GET /api/capacity/profiles`
