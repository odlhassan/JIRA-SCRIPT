# Calendar worklogs and Jira sign-in

Atlassian sign-in now also gates the whole EPR application. New users need Admin approval; Member/Admin roles, eight-hour EPR sessions, access administration, and private audit records are described in [EPR authentication](EPR_AUTHENTICATION.md). The existing OAuth app and scopes are reused. Logout preserves approvals; Atlassian privacy erasure removes access/profile records and anonymizes audit IDs. Access records are included in periodic privacy reporting.

## Business Cases

Employees log time from their Availability calendar and personal Atlassian sign-in ensures Jira credits the actual user. Shared server credentials cannot impersonate another employee.

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
- The returned Jira ID is stored before a separate GET verifies author and duration. The current calendar updates in browser memory. Other sessions/reloads use the canonical snapshot until the next refresh. No canonical rows are fabricated and no schema changes are needed.
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

On `closed` or `updated`, the integration erases matching OAuth sessions and removes personal profile/comment fields from its request receipts. Worklog IDs and non-personal idempotency state remain, preventing duplicate writes. Sign-out erases the account's integration sessions and receipt profile fields; other sessions for the same account will need to sign in again. Server sessions expire after 30 days, and rejected refresh credentials trigger cleanup. This cleanup does not delete Jira worklogs or change the organization's canonical reporting database.

The reporter uses an existing authorized OAuth token, refreshes it server-side when needed, and sends no credentials to logs. Persistent files remain under `/home/data/worklog_requests/oauth`. Documentation: [Atlassian user privacy guide](https://developer.atlassian.com/cloud/jira/platform/user-privacy-developer-guide/).

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
