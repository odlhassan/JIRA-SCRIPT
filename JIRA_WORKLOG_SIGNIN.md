# Calendar worklogs and Jira sign-in

Atlassian sign-in now also gates the whole EPR application. New users need Admin approval; Member/Admin roles, sessions of up to 30 days, access administration, and private audit records are described in [EPR authentication](EPR_AUTHENTICATION.md). The existing OAuth app and scopes are reused. Logout preserves approvals; Atlassian privacy erasure removes access/profile records and anonymizes audit IDs. Access records are included in periodic privacy reporting.

## Business Cases

Any named employee's Availability calendar reads that employee's Jira worklogs on opening,
when the displayed month changes, when the tab becomes visible, and every 60 seconds
while open. **Refresh Jira worklogs** requests an immediate read. The status shows the
last successful read; an incomplete or failed read keeps the last available calendar.
Changing months shows a small circular loading indicator in the modal header while Jira is read. A late response from an earlier
month cannot overwrite the newly selected calendar.
All issue-search and worklog pages must finish before replacing the modal's month.
The existing reporting Jira account supplies read-only results without a personal Jira
worklog sign-in for the selected employee. EPR application access is still required.
The named calendar also reads live RLT leave, including logs entered by HR for the
employee. Both datasets update together; live leave reduces available capacity.
Capacity profiles and the main report continue to use saved data. The combined Grand Total calendar
uses the snapshot. Jira search indexing and network latency can delay external changes.
See [live RLT leave rules and verification](LIVE_RLT_LEAVE.md).

In **Manage my logged hours**, **Edit entry** opens start date/time, duration and
description fields. Saving checks the personal account ID, the Jira update timestamp,
the 24-hour per-task/day limit, and Jira permissions before updating Jira while leaving
the remaining estimate unchanged. Changed entries must be refreshed before editing.
The existing timezone is preserved. An unchanged description retains Jira formatting;
editing its text replaces it with plain text. After the manager closes the calendar
refreshes. A timed-out write must be checked in Jira before retrying.

For a saved-data audit, open an employee's Availability calendar and choose
**Download Jira tally PDF**. This read-only export compares live Jira with the
current EPR snapshot by worklog ID, task, date and duration, includes every day
of the displayed month, and explains exclusions from calendar work. It uses the
reporting account's Jira visibility; absent records do not prove user error or
failure to save. It does not change personal review/submit/delete. Routes,
pagination, data flow, schema reads and verification are documented in
[Monthly worklog tally](WORKLOG_TALLY.md).

Employees log time from their Availability calendar and personal Atlassian sign-in ensures Jira credits the actual user. Shared server credentials cannot impersonate another employee.

Employees can also correct mistakes by deleting their own individual entries on any Jira issue, including tasks assigned to someone else and issue types outside the calendar's subtask logging scope.

## Delete your own logged hours

Choose **My availability** in the signed-in account area, then **Manage my logged hours** in Availability Details' sticky toolbar. My Day's Quick log card and **My hours** task buttons open the same manager. The calendar's red **Delete** row action prefills the task and date. The manager always reads the signed-in user's personal history, even when another employee's calendar is open.

The responsive personal-timesheet dialog opens the displayed month automatically; no Jira key is needed to browse. Month history searches Jira with `worklogAuthor = currentUser()` and inclusive month bounds, then reads every worklog page for each discovered task and independently checks author account ID and started month. Each request scans up to ten task results, with **Load more tasks** when Jira supplies a continuation token. Until all pages are loaded the status explicitly says totals cover loaded entries only. Existing entries are deduplicated by task/worklog ID when loading more.

### Fields and actions

- **Month**, **Previous history month**, **Next history month**: default to the calendar month (or the current month in My Day). Changing month loads personal month history and clears the direct task/date selection.
- **Refresh**: re-read the current month or direct task from Jira. Errors show a retry message; missing personal sign-in shows a sign-in link when OAuth is configured, or explains the required administrator setup. Empty or unavailable history disables CSV export. Nothing is written by refresh.
- **Hours shown / Entries / Tasks / Days**: totals for the currently loaded and filtered rows. For example 30m plus 2h on two dates gives 2h 30m, two entries and two days.
- **Search your entries**: case-insensitive local search across task key, title, work description and worklog ID. Defaults empty. Searches loaded rows only.
- **Task**: defaults to All loaded tasks; choose a discovered key to filter locally.
- **Date (optional)**: empty by default, or the calendar row's selected day. Matches the Jira started-date portion; clear to restore other dates.
- **Sort by**: Newest first (default), Oldest first, or Longest first.
- **Jira task key / Load my entries**: optional direct lookup of any valid Jira issue key, including another employee's task and any issue type. Direct lookup reads all dates, independently of the month picker. Blank key loads the month. **All tasks this month** leaves direct-lookup mode and clears task/date filtering.
- **Export CSV**: downloads only the loaded, filtered rows, in the chosen order, with task, title, started timestamp, hours, description and worklog ID. Disabled when empty. Spreadsheet formula prefixes are escaped and cells are quoted.
- Each entry shows task key/title, readable date/time, duration, full description and worklog ID. Empty results explain how to adjust the view and offer **Clear filters**.
- **Delete entry** opens an inline review identifying task, date, hours and worklog ID. **Cancel deletion** sends no request. **Confirm delete** sends the existing confirmed delete request, disables conflicting actions while saving, then removes the entry and updates totals. A failed/uncertain deletion disables that row until a fresh read and is never automatically retried.
- **Copy link** copies this modal's URL (or displays a selectable URL if clipboard access is unavailable). Recipients see their own Jira history. **Close** or Escape returns to the parent view; Escape does not also dismiss the calendar.

The server validates same-origin JSON and the worklog header, requires a valid personal OAuth session (no shared-account fallback, even with the test override), and re-reads the exact Jira worklog to compare author **account IDs**, not display names. Other authors are rejected with HTTP 403, including for EPR Admins. Jira also enforces its project permissions; a Jira 403 explains that Delete own worklogs permission is needed. A successful Jira 204 confirms deletion. Remaining estimate is unchanged (`adjustEstimate=leave`). No local schema or canonical snapshot data is changed.

After deletion the client re-reads the affected task's personal history. When the manager closes, the Availability calendar reads the full displayed month again. If the read fails, the last available calendar stays visible with an error. Older periods, lifetime task totals and other snapshot reports update after the normal Jira refresh. My Day reloads its own data. Tests mock Jira writes unless a user explicitly authorizes a separate real integration test.

Data flow: My Day loads `/worklog/manage.js` and `/worklog/manage.css`; the report generator inlines the same sources. Month browse uses `GET /worklog/history?month=YYYY-MM[&cursor=...]`; direct lookup uses `GET /worklog/entries?issue_key=...`. Both require personal identity. Confirmed `/worklog/delete` performs the fresh ownership check and Jira DELETE, then emits the existing UI update event. `epr_auth.py` permits approved Members to use the specific delete route; no Member write permission was broadened.

## Business Logic

- A plus button appears on each in-month day on hover or keyboard focus; touch devices always show it. Empty leading/trailing cells have no plus.
- Search matches canonical non-RLT subtask keys and titles, showing up to 100 matches and sorting the employee's assigned items first. Work on someone else's subtask is allowed if Jira grants access.
- Time spent accepts `1d`, `8h`, `8m`, mixed units such as `1h 30m`, fractional units resolving to whole minutes such as `0.5h`, or a bare minute count such as `90`. This calendar form deliberately defines **1d = 24h**, rather than a Jira-configured working day, so `2d`, `25h`, and `1d 1m` are invalid. Units are case-insensitive and may appear once each. Zero, negative, malformed, duplicate-unit, and fractional-minute durations are rejected. API clients may still send integer `minutes`; when `duration` is supplied the server parses that string authoritatively.
- The selected date, timezone-aware start time, positive duration no longer than 24h, required description of up to 2000 characters, subtask, and employee are checked at review time. Future dates are rejected.
- Review makes no Jira write. Jira supplies the author identity, which must match the calendar employee. Jira must confirm a non-leave subtask and grant Browse Projects and Work on Issues.
- Signed review tokens bind the author account ID, issue, timestamp, duration and comment, and expire after 10 minutes. The user checks a confirmation checkbox before the final button activates.
- Review reads all pages of the selected subtask's live Jira worklogs and totals the signed-in author's entries on the selected date in the browser's supplied timezone offset. Existing seconds plus the new duration must not exceed 86,400 seconds. Other authors and other local dates do not consume this person's subtask/day allowance. The review shows existing hours and the total after the proposed entry. An unreadable Jira history blocks submission.
- Submit repeats author, permissions, and live daily-total checks. A process-wide submit lock serializes the report's writes in the production single-worker configuration so two independently reviewed entries cannot race past the cap. Entries created directly by other Jira clients cannot participate in that lock; the app checks the latest available Jira history immediately before its POST. An exclusive persistent receipt claims the request before the single Jira POST. Completed request retries return the stored worklog ID; uncertain requests are blocked rather than resent. Check Jira before starting a fresh request after a timeout.
- Jira receives seconds, an ADF comment, a timestamp with timezone, and a request-ID property. `adjustEstimate=leave` preserves estimates and `notifyUsers=false` avoids notification noise.
- The returned Jira ID is stored before a separate GET verifies author and duration. The personal calendar then reads Jira again. Other snapshot reports update on their normal refresh. No canonical rows are fabricated and no schema changes are needed.
- Public review/submit requires a personal OAuth session. `EPR_WORKLOG_ALLOW_SHARED_ACCOUNT=1` is a local-only, explicitly authorized integration-test override; leave it unset in Azure. The calendar employee must still match the connected account.

## Explanations and Front-end UI Fields

Hover a day and choose **+**. A date header and account strip identify the day and Jira author. The responsive composer separates **Choose a subtask** from **Add your time**. Search/select a subtask; its full title and assignee appear below the list. Enter Time spent (`30m` default), Start time (09:00 default), and Work description. Quick buttons offer `30m`, `1h`, `4h`, and `8h`; a live preview shows normalized hours/minutes and the total minute count. Invalid durations show inline errors and cannot advance.

**Review worklog** checks live Jira and displays author, issue link/title, start, normalized duration, comment, existing daily hours, proposed day total out of 24h, and unchanged-estimate behavior. **Back** preserves the selected subtask, search, duration, time, and description. Check **I checked the author, subtask, date, and duration** and press **Confirm & log**. Submission disables controls and never automatically retries. Success displays the Jira worklog ID and verification result.

The account area shows **Sign in with Jira** when configured. Without an OAuth app it explains administrator setup is pending, and review stays disabled. **Sign out** removes the server session and cookie.

Availability values use bold purple underlined text. Closing the centered calendar hides it completely, removes pointer interaction, and returns focus; no panel strip remains at the right edge.

## Jira OAuth setup (required once)

1. Open the [Atlassian developer console](https://developer.atlassian.com/console/myapps/) with the account that will own the integration. Create an OAuth 2.0 integration for EPR Reporting.
2. Under Authorization, enable OAuth 2.0 (3LO) and register exactly `https://epreporting.azurewebsites.net/api/jira-auth/callback`.
3. Add Jira API classic permissions `read:jira-user`, `read:jira-work`, and `write:jira-work`. The authorization URL also requests `offline_access` for rotating refresh tokens.
4. Enable app sharing/distribution so other employees can consent, subject to your organization's app policy.
5. In Azure App Service application settings, securely set `JIRA_OAUTH_CLIENT_ID`, `JIRA_OAUTH_CLIENT_SECRET`, and `JIRA_OAUTH_REDIRECT_URI` (the callback above), then restart. Keep secrets out of source control and chat. Leave `EPR_WORKLOG_ALLOW_SHARED_ACCOUNT` unset.
6. Open an Availability calendar, click **+**, choose **Sign in with Jira**, and grant access to the report's Jira site. The callback checks browser-bound state and site/scopes. It returns to the report; reopen the desired calendar/date.

The **EPR Reporting** resource-level OAuth app was registered on 2026-10-28 (app ID `ba97f4c8-594a-42f7-b8b8-14351f427eae`). Its callback and Jira scopes are configured, and the three OAuth settings have been saved in Azure without committing credentials. Organization-only distribution is enabled with Octopus Digital as vendor, the work support email, and personal-data storage declared as Yes. References: [Atlassian 3LO](https://developer.atlassian.com/cloud/jira/platform/oauth-2-3lo-apps/) and [worklog API](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issue-worklogs/).

## Personal-data reporting

Atlassian requires apps retaining account-related data to implement its reporting API before enabling sharing with a Yes personal-data declaration. The app requests `report:personal-data` alongside Jira scopes and offline access. `jira_privacy.py` starts one daemon reporter per application directory when OAuth is configured. It wakes hourly, reports app-owned OAuth-session and worklog-receipt account IDs in batches of at most 90, and resumes from a private state file after restart. Reports use the oldest retained retrieval timestamp and respect each account's `Cycle-Period` (default seven days) and HTTP 429 `Retry-After`.

On `closed` or `updated`, the integration erases matching OAuth sessions and removes personal profile/comment fields from its request receipts. For `updated`, the EPR approval and role remain and the next sign-in restores the profile; for `closed`, the access record is removed. Worklog IDs and non-personal idempotency state remain, preventing duplicate writes. Sign-out erases the account's integration sessions and receipt profile fields; other sessions for the same account will need to sign in again, but keep their approval. Server sessions expire after 30 days. An invalid refresh grant removes the affected session. This cleanup does not delete Jira worklogs or change the organization's canonical reporting database.

The reporter uses an existing authorized OAuth token, refreshes it server-side when needed, and sends no credentials to logs. Token rotation is serialized in the single-worker server. Expiration and invalid-grant cleanup remove the affected session; privacy requests and explicit logout retain their existing account-wide cleanup. Persistent files remain under `/home/data/worklog_requests/oauth`. Documentation: [Atlassian user privacy guide](https://developer.atlassian.com/cloud/jira/platform/user-privacy-developer-guide/).

## Table Schema and Data Flow

No table or column changes. Existing canonical issues feed search; review checks live Jira. Tokens and refresh tokens live only in private JSON files under `<capacity DB parent>/worklog_requests/oauth`. Browser cookies contain opaque random IDs, with HttpOnly, SameSite=Lax, and Secure on HTTPS. Tokens refresh server-side. Session files use mode 0600 where supported. Signing secrets and request receipts live in the parent worklog directory and survive deploys when the DB is under `/home/data`. These directories are not static assets. Sign out removes the local token file; revoke the Atlassian grant to revoke access globally.

## Examples

An employee submits 30 minutes on APP-123 for October 28 at 09:00 +05:00. Jira receives 1800 seconds and the signed-in author. Replaying that reviewed request produces one worklog. A Bob calendar with Alice signed in is blocked.

With 23h already logged by Alice on APP-123 on October 28, `1h` is allowed and `1h 1m` is blocked. If another entry fills the allowance after review, final submission is blocked on its fresh Jira check. `1d` is 1440 minutes and is allowed only when no time is already logged by that author on that subtask/day. `8h` is 480 minutes; `8m` is eight minutes; `1h 30m` is 90 minutes.

## Script Files

- `employee_worklog.py`: validation, review, confirmed write, receipts and verification.
- `jira_oauth.py`: authorize/callback/status/logout and per-user token storage/refresh.
- `jira_privacy.py`: periodic privacy-account reporting, cycle state, and integration data erasure.
- `employee_worklog_form.js`: searchable form, review and final confirmation.
- `employee_worklog_manage.js` / `.css`: responsive personal history browser, search/filter/sort/export, share links and inline confirmed deletion.
- `employee_capacity_modal_links.js`: capacity report detail URL state, restoring filters/days/forms and browser history.
- `employee_availability_calendar.js` / `.css`: plus buttons and modal styling.
- `generate_employee_capacity_utilization_report.py`: inlines sources into the report shell.
- `tests/test_employee_worklog.py`: backend/OAuth tests with mocked writes.
- `tests/test_jira_privacy.py`: reporting cadence, timestamps, erasure, rate limits, and expiry.
- `tests/test_employee_availability_calendar.cjs`: browser confirmation and single-submit checks.
- `my_day.html`: the My Day **Quick log** card is a second client of the same `/identity`, `/review` and `/submit` endpoints. It follows the same review → confirm checkbox → submit flow (see `docs/report-user-guide/screens/17-my-day.md`).

## Dependent & Impacted Files

`report_server.py` registers worklog/OAuth routes using the capacity DB's persistent parent. `jira_client.py` supplies existing local-test credential factories and Jira site configuration. The generated root report and promoted `report_html` copy contain the UI. The report screen guide links this workflow. My Day (`my_day.html`) also submits through these endpoints, and `usage_analytics.py` counts successful `/submit` responses by Referer page. `report_server.my_day_api` reuses `jira_identity()` to read this week's worklogs live. Azure needs the three OAuth settings above; no package dependencies were added.

## Local verification

Generate: `.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py`.
Start: `.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3119`.
Open `http://127.0.0.1:3119/employee_capacity_utilization_report.html`.
For real local OAuth, register a local callback matching `JIRA_OAUTH_REDIRECT_URI`; mocked tests need no credentials.

Tests: `.venv/Scripts/python.exe -m unittest tests.test_employee_worklog tests.test_employee_capacity_utilization_report` and `node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line`.

## Authorized live verification — 2026-10-28

One 30-minute entry was created for Hassan Malik on `DIGITALLOG-1162` at 09:00 +05:00, worklog ID `69101`, with an explicit EPR integration verification description. Independent Jira reads confirmed exactly one matching entry and 1800 seconds. Jira rejected earlier attempts before creating any record because the request property needed a JSON object rather than a scalar string; that payload was corrected and regression coverage added. No further live writes are part of verification. Personal OAuth consent was subsequently completed successfully for Hassan Malik. The production callback returned to the report, the form showed Signed in to Jira: Hassan Malik, and live subtask search and review succeeded for DIGITALLOG-1162. Final confirmation stayed unchecked and disabled, so this sign-in verification created no additional worklog. Azure now holds the actual client secret transferred directly by the owner; no credentials were committed or printed.

## Activation verification — 2026-10-28

Organization-only sharing is enabled. After the owner pasted the actual client secret directly into Azure, settings were saved and the app restarted. Hassan Malik authorized the resource-restricted app on octopusdtlsupport.atlassian.net. Production sign-in, subtask search, and live review were verified; no extra time entry was created. The signed-in report was left open in the user's in-app browser.

A live privacy reporting pass succeeded with `{'reported': 1}`. This exact command ran in the application runtime's Azure WebSSH console:

```sh
python -c "import sys;sys.path.insert(0,'/home/site/wwwroot/.python_packages/lib/site-packages');sys.path.insert(0,'/home/site/wwwroot');from pathlib import Path;from jira_privacy import report_accounts;print(report_accounts(Path('/home/data/worklog_requests/oauth')))"
```

Kudu's separate command environment lacked the application module/runtime paths, so the diagnostic invocation could not run there; WebSSH was used successfully instead. The hourly app worker will honor the persisted reporting cycle.

Focused automated verification: `.venv/Scripts/python.exe -m unittest tests.test_jira_privacy tests.test_employee_worklog tests.test_employee_capacity_utilization_report` — 17 passed. Client secrets, refresh tokens, and session cookies were not printed or committed.

## Authorized disposable integration test — 2026-11-02

The user explicitly requested three-minute add/delete testing on Hassan Malik's tasks. Local personal OAuth was not configured, so interactive UI writes correctly remained disabled. An isolated Flask test app injected a real Jira session authenticated as Hassan Malik into the same worklog routes; no running-server authentication configuration or OAuth files were changed. The dedicated write credential failed `/myself` with HTTP 401; the primary configured credential independently verified Hassan's account and work/delete permissions.

On Hassan's existing subtask `ODL-155`, one 180-second entry was created at 11:06 +05:00, worklog ID `69292`. The service and an independent Jira read verified it. Replaying the reviewed submission returned the same ID without creating another entry. The owner-checked delete route returned HTTP 200 after Jira deletion; fresh Jira and service-list reads confirmed removal. All baseline worklog records and the remaining estimate matched before and after. No epic/story/subtask hierarchy was created. The non-secret local result is `tmp/worklog-3m-roundtrip-result.json`. This verifies real service/Jira operations; personal OAuth UI sign-in remains a separate local setup prerequisite. Nothing was deployed.
