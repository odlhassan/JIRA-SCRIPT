# Monthly Jira / EPR worklog tally

## Business Cases

Employees and administrators check whether saved Jira worklogs reached EPR
without changing employee hours. Open one employee's Availability calendar,
choose the month, then **Download Jira tally PDF**. Normal EPR sign-in and
approval are required in production. Approved Members already have organization-
wide report read access. The tally uses the configured reporting Jira account's
visibility; it cannot impersonate the selected employee.

## Business Logic

- A read-only SQLite transaction captures the same issue run
  (`canonical_refresh_state.last_success_run_id`) and worklog run (the last
  persisted `canonical_worklogs` row's run) as the report's existing readers.
  It reads the employee's selected-month authored entries, preserving IDs, and
  closes the database before making network requests.
- Historical snapshot authors are matched by normalized display name. Exactly
  one matching Jira display name must resolve to an account. Live worklogs are
  independently filtered by that account ID; ambiguous names fail.
- Live issue discovery searches all projects and types, with a one-day search
  margin around the month for Jira profile/timezone differences. Exact month
  inclusion uses the saved worklog `started` date portion, matching the report.
  Snapshot tasks are read directly too, so entries moved outside the month can
  be found. Every issue/worklog page is read and worklogs deduplicate by ID.
- IDs receive **Match**, **Only in Jira**, **Not returned by Jira**, or
  **Different**. A match requires equal task, saved date and duration; equal
  totals alone are insufficient. A moved-out-of-month live row stays in details
  but does not add to the live month total. Every calendar day is included,
  including **No entries in either source** days.
- Source comparisons include all types and RLT leave worklogs. The separate
  **EPR calendar work** column applies the existing non-RLT subtask rule,
  regardless of assignee. Story/leave/unknown-metadata exclusions are explained.
  Summary scope, teams and leave inclusion switches do not narrow the audit.
  It compares the current persisted snapshot, not in-memory browser overlays.
- Invalid/unreadable/incomplete Jira pages fail the export, rather than creating
  a partial PDF or zero-hour fallback. Bounds: 20 employee-search pages, 200
  issue-search pages, 1000 worklog pages per issue; 10s connect / 30s read timeout.
  Hitting a bound also fails explicitly.
- Jira reads reflect the reporting account's visible records and search index.
  Absence does not establish deletion, failure to save or user error. Concurrent
  Jira changes can affect a multi-request comparison. No worklogs, refreshes,
  schema, OAuth permissions or canonical database rows are modified. Generated
  JSON/PDF data is streamed without a persistent server cache.

## Examples

During the November 2 investigation, Faiza's October records initially matched:
29 entries / 147h. No returned worklogs existed on October 26, 28, 29 or 30.
Later Jira reads found changes made after that snapshot: worklog `69293` on
`O2-4524` was moved from November 2 to October 19 at 12:46 PKT, and `69300` on
`RLT-434` was created at 12:42 PKT for 8h leave on October 26. These now appear as
**Only in Jira** for October until EPR refreshes. October 28-30 still had no
returned records. These observations do not establish why the user expected logs.

An ID moved from October 26 / 1h to November 2 / 1h is **Different**. October's
live total excludes it, while the older October snapshot still includes it.

## Explanations

The PDF contains an overview, every day's Jira/EPR/calendar hours, then individual
worklog IDs with task/title, each source's date/hours, EPR planned task dates,
match status and calendar inclusion. It prints UTC generation time, Jira reader,
employee account ID, issue/worklog snapshot IDs and fetch completion time. Saved
worklog dates determine the month; task planned start/due dates do not.

## Front-end UI Fields

- **Month / Previous / Next:** existing displayed calendar selection; this
  employee/month is captured when requesting the tally.
- **Download Jira tally PDF:** available for one named employee, not Grand Total;
  disables while reading and restores itself afterward. Downloads an actual PDF.
- **Status:** announces progress, employee/month, completion or the failure reason.
  Authentication/non-PDF/error responses never become a purported PDF download.
- **Manage my logged hours:** retains personal live-history and owner-only
  deletion behavior; the tally does not change the review/submit/delete flow.

## Script Files

- `worklog_tally.py`: snapshot capture, paginated Jira reads, reconciliation, PDF.
- `report_server.py`: protected GET `/api/employee-capacity-utilization/tally`
  (JSON) and `/api/employee-capacity-utilization/tally.pdf` (download), both taking
  required `employee` and `month`; success has `Cache-Control: no-store, private`.
- `employee_availability_calendar.js` / `.css`: download/status/layout.
- `generate_employee_capacity_utilization_report.py`: inlines calendar sources
  into the root report shell, promoted to `report_html` on server startup.
- `requirements.txt`: ReportLab runtime dependency; `requirements-dev.txt` adds
  pypdf for PDF text assertions in tests.
- `tests/test_worklog_tally.py`: reconciliation, snapshots, paging, PDF and routes.
- `tests/test_employee_availability_calendar.cjs`: download/error/mobile checks.

## Dependent & Impacted Files

The personal worklog guide, report screen guide, README and expected-files guide
link this workflow. Existing Excel exports, Colossal Refresh, review/submit/delete,
OAuth and Member write permissions retain their behavior. Azure's existing deploy
workflow vendors ReportLab from `requirements.txt` under Python 3.11.

## Table Schema

No schema changes. Reads existing `canonical_refresh_state` (`id`,
`last_success_run_id`), `canonical_refresh_runs` (`run_id`, `ended_at_utc`),
`canonical_issues` (`run_id`, `issue_key`, `summary`, `issue_type`, `start_date`,
`due_date`) and `canonical_worklogs` (`run_id`, `worklog_id`, `issue_key`,
`worklog_author`, `started_date`, `hours_logged`). IDs/dates/names are TEXT,
hours are REAL. Existing primary keys are `(run_id, issue_key)` for issues and
`(run_id, worklog_id)` for logs. SQLite opens with `mode=ro` and is closed before
Jira reads. No new tables or columns are introduced.

## Data Flow

Calendar employee/month -> protected tally GET -> snapshot transaction ->
reporting Jira account and exact employee lookup -> paginated issue discovery
plus snapshot tasks -> account-ID-filtered worklogs -> compare by ID -> JSON/PDF
stream -> browser download. HTTP 400 rejects invalid selection, 409 rejects
unavailable/incomplete data, 502 rejects Jira failures, 503 rejects SQLite failures.

## Local verification

From `E:\JIRA SCRIPT`:

```powershell
.venv/Scripts/python.exe -m pip install -r requirements-dev.txt
.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py
.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3122
```

Open `http://127.0.0.1:3122/employee_capacity_utilization_report.html`, select an
employee/month and download the tally. Requires a successful local canonical
snapshot and configured Jira read credentials; production requires approved EPR
sign-in. Tests mock Jira and require no production credentials:

```powershell
.venv/Scripts/python.exe -m unittest tests.test_worklog_tally tests.test_employee_worklog tests.test_employee_capacity_utilization_report tests.test_epr_auth
node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line
```
