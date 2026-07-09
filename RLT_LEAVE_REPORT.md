# RLT Leave Intelligence Report

## Project and Window
- Project Key: `RLT`
- Project Name: `RnD Leave Tracker`
- Reporting Window: `2026-06-01` to `2026-08-31`

## Executive Summary
- Planned Taken (hours): `192.00`
- Unplanned Taken (hours): `64.00`
- Unknown Taken (hours): `252.00`
- Planned Not Yet Taken (hours): `40.00`
- Planned Not Yet Taken (No Entry count): `1`
- Defective subtasks listed: `129`
- Clubbed leave subtasks: `31`

## Assignee-wise Summary
| Assignee | Planned Taken (h) | Unplanned Taken (h) | Unknown Taken (h) | Planned Not Yet Taken (h) | No Entry Count | Unknown Count |
| --- | --- | --- | --- | --- | --- | --- |
| Aiza Hamid | 0.00 | 0.00 | 40.00 | 0.00 | 0 | 5 |
| Arsalan Zafar Khan | 0.00 | 0.00 | 0.00 | 16.00 | 0 | 2 |
| awais akhter | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 4 |
| DANIYAL AHMAD | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 2 |
| Faiq Butt | 0.00 | 0.00 | 24.00 | 0.00 | 0 | 3 |
| Faiza Nasir | 120.00 | 0.00 | 0.00 | 0.00 | 0 | 3 |
| Hamza Ali | 0.00 | 8.00 | 0.00 | 0.00 | 0 | 8 |
| Hassan Malik | 8.00 | 32.00 | 16.00 | 0.00 | 0 | 7 |
| Hassan Saeed Wattoo | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 25 |
| Hassan Wali | 0.00 | 16.00 | 0.00 | 0.00 | 0 | 0 |
| Ibrahim Ahmed Qureshi | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 6 |
| Maria Sharafat | 0.00 | 0.00 | 0.00 | 0.00 | 1 | 2 |
| Muhammad Abbas | 16.00 | 0.00 | 0.00 | 0.00 | 0 | 4 |
| Muhammad Abdul Wasi | 0.00 | 0.00 | 4.00 | 0.00 | 0 | 1 |
| Muhammad Abdullah | 0.00 | 0.00 | 16.00 | 0.00 | 0 | 6 |
| Muhammad Ahmad Saleem | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 3 |
| Muhammad Imran Aslam | 8.00 | 0.00 | 0.00 | 24.00 | 0 | 0 |
| Muhammad Usman Javed | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 2 |
| Muhammad Zeeshan Aslam | 0.00 | 0.00 | 136.00 | 0.00 | 0 | 1 |
| Namra Zahid | 16.00 | 0.00 | 8.00 | 0.00 | 0 | 2 |
| Sarmad Sabir | 16.00 | 0.00 | 8.00 | 0.00 | 0 | 10 |
| Syed Yousaf Qadri | 0.00 | 0.00 | 0.00 | 0.00 | 0 | 7 |
| Taimur Zahid | 8.00 | 0.00 | 0.00 | 0.00 | 0 | 7 |
| Zeeshan Sarwar | 0.00 | 8.00 | 0.00 | 0.00 | 0 | 18 |

## Defective and No Entry
- `No Entry` means planned leave subtask is missing planned date and/or original estimate while no hours are logged.
- Unknown classification subtasks are shown separately and are not merged into planned/unplanned totals.

## Clubbed Leave
- Clubbed leave means one subtask represents more than one day (for example logged/estimated hours > 8 or multi-day date span).

## Data-Quality Notes
- Month/week forecasts use Jira date fields only.
- Subtasks without Jira dates are not bucketed into week/month and are reported as data-quality issues.
- Hours are primary; days are derived by date-aware hours/day (Ramadan dates use Ramadan hours/day; other dates use standard hours/day).

## Employee Performance Scoring Note
- Employee Performance uses weighted normalized Advanced Score factors instead of raw penalty multipliers. RLT unplanned leave hours feed the `Leave Reliability` factor as `100 × (1 − min(1, unplanned_leave_hours / employee_capacity_hours))`, then that factor contributes by its configured weight.