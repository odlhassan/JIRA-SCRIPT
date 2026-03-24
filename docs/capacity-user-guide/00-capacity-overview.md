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
