# Incremental Availability calendar reads

## Purpose and behavior

The named employee calendar shares a bounded, process-local reporting-account cache
across approved EPR viewers. First access reads complete worklogs and live RLT leave.
Reopening the same employee/month within 30 seconds reuses the successful result with
zero Jira requests. After that, Jira discovery rechecks issue membership and metadata;
unchanged issue worklog pages are reused. Explicit Refresh bypasses the result's
30-second reuse window but still reuses unchanged worklog pages.

## Business rules and data flow

`jira_calendar_cache.py` wraps the existing reporting session, preserving the established
worklog tally and RLT computations. It is scoped by database, Jira site and a one-way
credential fingerprint. Cached data is never returned before the existing EPR access
gate. Personal OAuth credentials are not used by this shared cache.

Jira updated/deleted worklog feeds are checked at most every 30 seconds across viewers.
Known changed IDs invalidate only that issue's cached pages. Newly seen changed IDs
are resolved through the bulk worklog read endpoint to find their issue. Changes in
issue-discovery fields also invalidate affected pages, covering RLT date, classification,
assignment and estimate changes. Every calendar recomputes from the current discovered
membership, so moves between months and leave reassignment affect both old and new views
as they are checked. A worklog issue page expires after 15 minutes for full reconciliation.
Employee account lookup and reporting identity are cached for one hour; a unique Atlassian
product account is distinguished from a same-name portal customer. Duplicate product
accounts still fail rather than guess.

Feed pagination must complete; its checkpoint uses a two-minute overlap because Jira
omits the last minute's changes. A timestamp ledger avoids repeated invalidation of the
same overlap events. External changes therefore have indexing/feed delay. Successful
EPR create/edit/delete invalidates that issue and the result reuse timers immediately.
No mutation request is retried by this cache.

A lock coalesces concurrent reads under the existing single Gunicorn worker. Cache state,
checkpoints and calendar results are committed together only after the full worklog and
RLT read succeeds. A failed read retains prior state. A restart clears the cache. Memory
limits bound pages, identities, worklog mappings and calendar results; hitting them resets
the cache for a fresh next read. This is not a durable second reporting database.

## UI and examples

The small modal header spinner remains. Status distinguishes **Updated at** from
**Checked at · No changes**. An unchanged result preserves the current day detail.
The JSON result includes `sync.mode`, `checked_at`, `changed_at`, and `network_reads`
for verification; cached responses report zero network reads. A 60-second unchanged
poll checks discovery/feed data and does not reread full worklog pages. PDF tally
downloads always use fresh direct reads rather than the shared calendar cache.

## Files and verification

`report_server.py` uses the cache only for the JSON calendar route;
`employee_worklog.py` invalidates after successful mutations;
`employee_availability_calendar.js` displays freshness;
`worklog_tally.py` and `live_rlt_leave.py` remain calculation owners.
`tests/test_jira_calendar_cache.py` covers reuse, concurrency, updates, deletions,
new IDs, metadata changes, failures and periodic reconciliation.

No SQLite schema or data changes. Test with:
`.venv/Scripts/python.exe -m unittest tests.test_jira_calendar_cache tests.test_worklog_tally tests.test_live_rlt_leave tests.test_employee_worklog`.
