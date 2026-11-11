# Live RLT leave in Availability Details

Calendar reads now share the [incremental Jira cache](JIRA_CALENDAR_CACHE.md).
RLT membership and metadata are rechecked, while unchanged worklog pages are reused;
classification and date changes still recompute the live daily values.

## Business cases and explanation

An employee's calendar must distinguish leave from missing worklogs. Namra Zahid's
October 19–22, 2026 RLT subtasks contain 8h each: the calendar must show 32h leave
and subtract that from the configured month's capacity. This is read-only and
available to approved EPR report viewers through the reporting Jira account.

## Business logic and data flow

The JSON tally route reads worklogs and live RLT leave before returning either to
the calendar. `live_rlt_leave.py` searches all pages of RLT issues assigned to the
resolved employee account ID, keeps subtasks only, and reads every worklog page.
Leave belongs to the RLT subtask assignee, even when HR authored its worklog.
Parent task accumulated time is not counted again. No issue-created date or month
filter narrows discovery: multi-day leave can cross months or be logged later.

The existing RLT normalization, summary-date fallback, continuous-leave redistribution
and daily aggregation functions from `generate_rlt_leave_report.py` supply the dates
and planned/unplanned classification. Only daily rows in the selected month return.
Logged leave with a missing classification counts as unclassified taken leave;
planned estimates without logged leave are not counted as taken. No hours are inferred
from missing worklogs. Invalid or incomplete reads fail explicitly rather than return
zero. Live worklogs and leave replace the modal's values together after success.

## Front-end fields

The small header spinner covers both Jira worklogs and RLT leave reads. **Leave taken**
and **Available capacity** show **Checking RLT…** before the first complete read, and
**Unavailable** on failure. A previously successful read remains visible with a stale
status on a later failure. Red daily badges open employee/date, RLT links, planned,
unplanned, unclassified and total taken hours. Available capacity is the configured
monthly capacity less live taken leave, floored at zero.

## Related code, dependencies and schema

- `live_rlt_leave.py`: paginated read and existing RLT computation reuse.
- `report_server.py`: JSON tally joins both live datasets; PDF audit stays unchanged.
- `employee_availability_calendar.js` and the generated report: modal rendering.
- `tests/test_live_rlt_leave.py`, `tests/test_worklog_tally.py` and the Availability
  browser tests cover assignment, dates, pagination, totals and failures.

No schema changes or DB writes. Existing canonical data is still used by the main
report and Grand Total calendar; the named employee modal has its own live overlay.
Personal Jira authentication remains mandatory for create/edit/delete worklog actions.

## Create and record leave from Availability

The user-authorized RLT path in `employee_subtask.py` creates a leave subtask under an existing RLT Task/Story tracker without an epic, with validated Planned/Unplanned Leave metadata, dates and personal assignment. `employee_worklog.py` accepts `entry_kind=leave` only with personal OAuth and an RLT subtask assigned to that exact account; review and submit both recheck it. The UI keeps leave recording separate from ordinary work. Creating the issue or its estimate never counts as leave taken; the user must explicitly review and confirm recorded leave time.

Dependent files/data flow: `employee_subtask_form.js` → `employee_worklog_form.js` → Jira issue/worklog → cache invalidation → live calendar read. New leave uses the configured RLT start field when available on create, otherwise Start date (`customfield_10015`). The reader requests both fields; shared `_normalize_subtasks` prefers the historical field and falls back to `customfield_10015` only when missing, preserving existing dates. `generate_rlt_leave_report.py` uses the same fallback after full fetch and in direct Jira mode. Classification, assignee ownership and continuous-leave redistribution are unchanged. See [subtask creation](JIRA_SUBTASK_CREATION.md).
