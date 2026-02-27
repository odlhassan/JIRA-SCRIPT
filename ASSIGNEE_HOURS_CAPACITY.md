# Assignee Hours Capacity Backend

This report now supports a SQLite-backed Capacity Planning form in the HTML report.

## User Guide Docs

Detailed functional documentation is available at:

- Module overview: `docs/capacity-user-guide/00-capacity-overview.md`
- Capacity settings page fields and behavior: `docs/capacity-user-guide/screens/01-capacity-settings-page.md`
- Assignee Hours integration and KPI logic: `docs/capacity-user-guide/screens/02-assignee-hours-capacity-integration.md`
- Nested View linkage and profile application flow: `docs/capacity-user-guide/screens/03-nested-view-capacity-integration.md`
- Unified cross-report info-drawer and logic docs: `docs/report-user-guide/00-report-overview.md`

## Files Updated

- Python generator and server:
  - `generate_assignee_hours_report.py`
- HTML outputs:
  - `assignee_hours_report.html`
  - `report_html/assignee_hours_report.html`
- Excel output:
  - `assignee_hours_report.xlsx`
- SQLite storage:
  - `assignee_hours_capacity.db`
  - `assignee_hours_capacity.sqlite`

## Run Modes

- Static generation:
  - `python generate_assignee_hours_report.py`
  - The generated HTML now embeds leave daily rows from `rlt_leave_report.xlsx` and can compute leave KPIs without backend API.
- Server mode (required for persistent capacity save/load):
  - `python generate_assignee_hours_report.py --server --port 5000`
  - Open `http://localhost:5000`

## Capacity APIs

- `GET /api/capacity?from=YYYY-MM-DD&to=YYYY-MM-DD`
- `POST /api/capacity`
- `POST /api/capacity/calculate`
- `GET /api/capacity/profiles`

## Reuse Saved Capacity

- Capacity settings are saved per date range.
- In the Capacity Planning form, use **Reuse Saved Capacity** to select a previously saved range.
- Click **Apply Profile To Current Range** to load those settings into the current selected report range.
- Then click **Save Capacity** to persist that profile for the current range.

## Capacity Formula

`available_capacity_hours = employees * (non_ramadan_workdays * standard_hours_per_day + ramadan_workdays * ramadan_hours_per_day)`

Where:

- Working days are Monday-Friday in selected date range
- Holiday dates are excluded
- Ramadan is a contiguous start/end range if provided

## Capacity Planning KPIs

- `Total Capacity`:
  - Raw profile capacity (`available_capacity_hours`)
- `Leave Hours`:
  - `planned_taken_hours + planned_not_taken_hours + unplanned_taken_hours`
- `Remaining Capacity`:
  - `Total Capacity - Leave Hours`
- `Project Planned Hours`:
  - Existing planned-minus-leaves value (caption updated)
- `Project Actual Hours`:
  - Total logged project hours excluding project key `RLT` (case-insensitive)
- `Project Plan - Actual Hours`:
  - `Project Planned Hours - Project Actual Hours`

## Capacity Subtraction Card

`Capacity After Leaves = Available Capacity - Project Actual Hours - Leave Hours`

## API Leave Metrics

`/api/capacity`, `/api/capacity/calculate`, and `POST /api/capacity` include:

- `planned_taken_hours`
- `planned_not_taken_hours`
- `unplanned_taken_hours`
- `taken_hours`
- `not_yet_taken_hours`
- `taken_days`
- `not_yet_taken_days`
- `remaining_balance_hours`
- `remaining_balance_days`
