# RLT Leave Intelligence Report

## Project and Window
- Project Key: `RLT`
- Project Name: `RnD Leave Tracker`
- Reporting Window: `2026-02-01` to `2026-04-30`

## Executive Summary
- Planned Taken (hours): `0.00`
- Unplanned Taken (hours): `0.00`
- Planned Not Yet Taken (hours): `0.00`
- Planned Not Yet Taken (No Entry count): `0`
- Defective subtasks listed: `0`
- Clubbed leave subtasks: `0`

## Assignee-wise Summary
| Assignee | Planned Taken (h) | Unplanned Taken (h) | Planned Not Yet Taken (h) | No Entry Count | Unknown Count |
| --- | --- | --- | --- | --- | --- |

## Defective and No Entry
- `No Entry` means planned leave subtask is missing planned date and/or original estimate while no hours are logged.
- Unknown classification subtasks are excluded from planned/unplanned totals.

## Clubbed Leave
- Clubbed leave means one subtask represents more than one day (for example logged/estimated hours > 8 or multi-day date span).

## Data-Quality Notes
- Month/week forecasts use Jira date fields only.
- Subtasks without Jira dates are not bucketed into week/month and are reported as data-quality issues.
- Hours are primary; days are derived by date-aware hours/day (Ramadan dates use Ramadan hours/day; other dates use standard hours/day).