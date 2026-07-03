# Capacity Module Overview

## Purpose

The Capacity module defines reusable capacity profiles by date range and connects those profiles to report calculations.

It covers:
- Profile setup and maintenance on the Capacity Settings page.
- Capacity and leave KPIs in the Assignee Hours report.
- Profile application in the Nested View report.
- RnD leadership data story with capacity/workload/pending-demand visuals.

## Users

- Delivery/project managers planning team utilization.
- Reporting users validating capacity vs actual logged hours.

## Prerequisites

- Run reports via server mode so API endpoints are available.
- Capacity database file must be accessible:
  - Default: `assignee_hours_capacity.db`
  - Azure: `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH` should normally point to `/home/data/assignee_hours_capacity.db`; quoted values are normalized, parent folders are created at startup, and an unwritable configured path falls back to `$HOME/data/assignee_hours_capacity.db` with a server log warning.
- Leave workbook should exist for leave-adjusted metrics:
  - Default: `rlt_leave_report.xlsx`

## Navigation

- Capacity settings page:
  - `/settings/capacity`
  - Typo-safe redirect also supported: `/settings/capactiy` -> `/settings/capacity`
- Assignee Hours report:
  - `/assignee_hours_report.html`
- Nested View report:
  - `/nested_view_report.html`
- RnD Data Story report:
  - `/rnd_data_story.html`
- Seating Planner:
  - `/settings/seating-planner`
- Team Capacity Planner:
  - `/settings/team-capacity-planner`

## Data Model (Profile)

Each saved capacity profile is keyed by:
- `from_date`
- `to_date`

Main fields:
- `employee_count`
- `standard_hours_per_day`
- `ramadan_start_date`
- `ramadan_end_date`
- `ramadan_hours_per_day`
- `holiday_dates`

## Core Formula

`available_capacity_hours = employee_count * (non_ramadan_weekdays * standard_hours_per_day + ramadan_weekdays * ramadan_hours_per_day)`

Rules:
- Workdays are Monday to Friday.
- `holiday_dates` are excluded.
- Ramadan hours apply only within Ramadan start/end range.

## API Surface

- `GET /api/capacity?from=YYYY-MM-DD&to=YYYY-MM-DD`
- `POST /api/capacity`
- `DELETE /api/capacity?from=YYYY-MM-DD&to=YYYY-MM-DD`
- `POST /api/capacity/calculate`
- `GET /api/capacity/profiles`

## Linkage to Reports

- Assignee Hours:
  - Loads/saves profiles through capacity APIs.
  - Uses profile + leave metrics to show remaining capacity.
- Nested View:
  - Loads saved profiles and applies selected profile to current report filter range.
  - Shows a read-only calendar preview for the selected profile, including Ramadan, holidays, leave tags, and range summary chips.
  - Can reset to project totals.
- RnD Data Story:
  - Applies saved capacity profiles to the selected date range.
  - Computes six leadership KPIs for department `Research and Development (RnD)`.
- Seating Planner:
  - Maintains the office floor seating layout with team/product color modes.
  - Exports the visible seating plan through a single-page A4 landscape, vector print view for saving as PDF.
- Team Capacity Planner:
  - Shows per-resource capacity, leave, logged work, and planned work for a selected team/date range.
  - The resource planned-work bar uses only assigned canonical subtasks in the selected range. Epic and story estimates are ignored even when those work items are assigned to the same resource.
  - The Stats toggle switches resource values between days and hours without reloading data.
- Support Hour Bookings (`/settings/performance`):
  - Reuses a saved capacity profile to compute each Technical Support Team member's system capacity hours for an admin-selected month, then lets the admin set leave/booking hours and per-project percentage allocations. See `docs/capacity-user-guide/screens/07-support-hour-bookings.md`.

## Deployment Notes

- Azure ZIP deploys vendor `requirements.txt` into `.python_packages/lib/site-packages` and marks `startup.txt` executable before packaging so the read-only Run From Package mount does not need a runtime chmod. The script adds `.python_packages` to `PYTHONPATH` before Gunicorn imports `report_server.py`.
- Capacity startup now creates the SQLite parent directory before opening the DB, which keeps Capacity Settings, Team Capacity Planner, and Employee Performance from failing cold start when the persistent DB folder has not been created yet.
