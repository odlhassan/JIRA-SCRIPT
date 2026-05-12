# RMI Jira Gantt View Redesign

## Summary
Redesign `rmi_jira_gantt_report.html` → `Gantt View` so it visually follows the existing `RMI Estimation & Scheduling` layout: sticky table-style left columns, month-wise timeline distribution, epic names outside bars, and bars containing schedule effort values instead of Jira keys.

## Key Changes
- Update `generate_rmi_jira_gantt_html.py` Gantt rendering to use `DATA.rmi_schedule_records` as the source so Gantt rows align with `RMI Estimation & Scheduling` filters, product grouping, year selection, and unit toggle.
- Replace the current simple SVG-only Gantt with a table-like layout:
  - Left columns: `#`, `RMI` showing epic name + Jira key link, `Product`, `Status`, `Most likely`, `TK Approved`.
  - Right timeline area: 12 month columns for the selected year, matching the schedule table’s month headers and sticky layout.
  - Product group and subtotal rows should mirror the schedule table styling.
- Draw one bar per dated epic across the month timeline:
  - Epic name/key stays in the left RMI column, not inside the bar.
  - Start date appears outside the left side of the bar.
  - Due date appears outside the right side of the bar.
  - Inside the bar show `TK Approved` by default.
  - When `Diagnostics` is enabled, inside the bar show `TK Approved | Epic Jira Estimate`.
- Add `jira_original_estimate_seconds` to `rmi_schedule_records` so the Gantt bar can display the epic-level Jira original estimate without recalculating from the DOM.
- Keep existing product filter, search, hours/days unit toggle, Diagnostics toggle, and `Only Jira Populated Epics` behavior consistent across metric cards, schedule table, and Gantt view.
- Update Gantt CSS alongside the generator so the new view handles long epic names, narrow bars, month grid lines, sticky columns, horizontal scrolling, and no-overlap date labels.
- Update `docs/report-user-guide/screens/13-rmi-jira-gantt-report.md` to document the new Gantt behavior and Diagnostics display.

## Test Plan
- Update `tests/test_rmi_jira_gantt_report.py` to assert:
  - `rmi_schedule_records` includes `jira_original_estimate_seconds`.
  - Rendered HTML includes the new Gantt table/grid classes and bar label format hooks.
  - Diagnostics mode affects Gantt bar estimate display.
  - Existing report controls, schedule table, navigation registration, and server route still render.
- Run focused checks:
  - `python -m unittest tests.test_rmi_jira_gantt_report`
  - `python generate_rmi_jira_gantt_html.py`
- If browser verification is available after implementation, start the local server and inspect:
  - `http://127.0.0.1:3000/rmi_jira_gantt_report.html`
  - Click `Gantt View`
  - Confirm product sections, month columns, left epic column, start/due labels, and Diagnostics bar text.

## Assumptions
- “Epic Jira Estimate” means the existing epic-level `jira_original_estimate_seconds` field used by the `Epic Estimates` metric.
- “Epic name + key” means the visible left RMI column should show the epic/RMI title plus a Jira key link.
- The selected year in `RMI Estimation & Scheduling` should also drive the Gantt timeline year, so both views stay aligned.

## How to experience latest changes on live localhost
After implementation, run `python run_server.py --port 3000` from `E:\JIRA SCRIPT`, then open `http://127.0.0.1:3000/rmi_jira_gantt_report.html` and switch to `Gantt View`.

## How to test locally
After implementation, run `python -m unittest tests.test_rmi_jira_gantt_report` and `python generate_rmi_jira_gantt_html.py`.
