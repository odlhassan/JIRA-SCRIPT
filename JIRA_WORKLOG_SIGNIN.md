# Calendar worklogs and Jira sign-in

## Business Cases

Employees log time from their Availability calendar and personal Atlassian sign-in ensures Jira credits the actual user. Shared server credentials cannot impersonate another employee.

## Business Logic

- A plus button appears on each in-month day on hover or keyboard focus; touch devices always show it. Empty leading/trailing cells have no plus.
- Search matches canonical non-RLT subtask keys and titles, showing up to 100 matches and sorting the employee's assigned items first. Work on someone else's subtask is allowed if Jira grants access.
- The selected calendar date and browser-local start time, 1–1440 whole minutes, a required description of up to 2000 characters, selected subtask, and employee are checked at review time. Future dates are rejected.
- Review makes no Jira write. Jira supplies the author identity, which must match the calendar employee. Jira must confirm a non-leave subtask and grant Browse Projects and Work on Issues.
- Signed review tokens bind the author account ID, issue, timestamp, duration and comment, and expire after 10 minutes. The user checks a confirmation checkbox before the final button activates.
- Submit repeats author and permissions checks. An exclusive persistent receipt claims the request before the single Jira POST. Completed request retries return the stored worklog ID; uncertain requests are blocked rather than resent. Check Jira before starting a fresh request after a timeout.
- Jira receives seconds, an ADF comment, a timestamp with timezone, and a request-ID property. `adjustEstimate=leave` preserves estimates and `notifyUsers=false` avoids notification noise.
- The returned Jira ID is stored before a separate GET verifies author and duration. The current calendar updates in browser memory. Other sessions/reloads use the canonical snapshot until the next refresh. No canonical rows are fabricated and no schema changes are needed.
- Public review/submit requires a personal OAuth session. `EPR_WORKLOG_ALLOW_SHARED_ACCOUNT=1` is a local-only, explicitly authorized integration-test override; leave it unset in Azure. The calendar employee must still match the connected account.

## Explanations and Front-end UI Fields

Hover a day, choose **+**, search/select a subtask, enter Start time (09:00 default), Time spent in minutes (30 default), and Work description. **Review worklog** checks Jira and displays author, issue link/title, start, duration, comment, and unchanged-estimate behavior. **Back** returns to the form. Check **I checked the author, subtask, date, and duration** and press **Confirm & log**. Submission disables controls and never automatically retries. Success displays the Jira worklog ID and verification result.

The account area shows **Sign in with Jira** when configured. Without an OAuth app it explains administrator setup is pending, and review stays disabled. **Sign out** removes the server session and cookie.

Availability values use bold purple underlined text. Closing the centered calendar hides it completely, removes pointer interaction, and returns focus; no panel strip remains at the right edge.

## Jira OAuth setup (required once)

1. Open the [Atlassian developer console](https://developer.atlassian.com/console/myapps/) with the account that will own the integration. Create an OAuth 2.0 integration for EPR Reporting.
2. Under Authorization, enable OAuth 2.0 (3LO) and register exactly `https://epreporting.azurewebsites.net/api/jira-auth/callback`.
3. Add Jira API classic permissions `read:jira-user`, `read:jira-work`, and `write:jira-work`. The authorization URL also requests `offline_access` for rotating refresh tokens.
4. Enable app sharing/distribution so other employees can consent, subject to your organization's app policy.
5. In Azure App Service application settings, securely set `JIRA_OAUTH_CLIENT_ID`, `JIRA_OAUTH_CLIENT_SECRET`, and `JIRA_OAUTH_REDIRECT_URI` (the callback above), then restart. Keep secrets out of source control and chat. Leave `EPR_WORKLOG_ALLOW_SHARED_ACCOUNT` unset.
6. Open an Availability calendar, click **+**, choose **Sign in with Jira**, and grant access to the report's Jira site. The callback checks browser-bound state and site/scopes. It returns to the report; reopen the desired calendar/date.

The **EPR Reporting** resource-level OAuth app was registered on 2026-10-28 (app ID `ba97f4c8-594a-42f7-b8b8-14351f427eae`). Its callback and Jira scopes are configured, and the three OAuth settings have been saved in Azure without committing credentials. Organization-only distribution requires the personal-data declaration below. References: [Atlassian 3LO](https://developer.atlassian.com/cloud/jira/platform/oauth-2-3lo-apps/) and [worklog API](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issue-worklogs/).

## Personal-data reporting

Atlassian requires apps retaining account-related data to implement its reporting API before enabling sharing with a Yes personal-data declaration. The app requests `report:personal-data` alongside Jira scopes and offline access. `jira_privacy.py` starts one daemon reporter per application directory when OAuth is configured. It wakes hourly, reports app-owned OAuth-session and worklog-receipt account IDs in batches of at most 90, and resumes from a private state file after restart. Reports use the oldest retained retrieval timestamp and respect each account's `Cycle-Period` (default seven days) and HTTP 429 `Retry-After`.

On `closed` or `updated`, the integration erases matching OAuth sessions and removes personal profile/comment fields from its request receipts. Worklog IDs and non-personal idempotency state remain, preventing duplicate writes. Sign-out erases the account's integration sessions and receipt profile fields; other sessions for the same account will need to sign in again. Server sessions expire after 30 days, and rejected refresh credentials trigger cleanup. This cleanup does not delete Jira worklogs or change the organization's canonical reporting database.

The reporter uses an existing authorized OAuth token, refreshes it server-side when needed, and sends no credentials to logs. Persistent files remain under `/home/data/worklog_requests/oauth`. Documentation: [Atlassian user privacy guide](https://developer.atlassian.com/cloud/jira/platform/user-privacy-developer-guide/).

## Table Schema and Data Flow

No table or column changes. Existing canonical issues feed search; review checks live Jira. Tokens and refresh tokens live only in private JSON files under `<capacity DB parent>/worklog_requests/oauth`. Browser cookies contain opaque random IDs, with HttpOnly, SameSite=Lax, and Secure on HTTPS. Tokens refresh server-side. Session files use mode 0600 where supported. Signing secrets and request receipts live in the parent worklog directory and survive deploys when the DB is under `/home/data`. These directories are not static assets. Sign out removes the local token file; revoke the Atlassian grant to revoke access globally.

## Examples

An employee submits 30 minutes on APP-123 for October 28 at 09:00 +05:00. Jira receives 1800 seconds and the signed-in author. Replaying that reviewed request produces one worklog. A Bob calendar with Alice signed in is blocked.

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

## Dependent & Impacted Files

`report_server.py` registers worklog/OAuth routes using the capacity DB's persistent parent. `jira_client.py` supplies existing local-test credential factories and Jira site configuration. The generated root report and promoted `report_html` copy contain the UI. The report screen guide links this workflow. Azure needs the three OAuth settings above; no package dependencies were added.

## Local verification

Generate: `.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py`.
Start: `.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3119`.
Open `http://127.0.0.1:3119/employee_capacity_utilization_report.html`.
For real local OAuth, register a local callback matching `JIRA_OAUTH_REDIRECT_URI`; mocked tests need no credentials.

Tests: `.venv/Scripts/python.exe -m unittest tests.test_employee_worklog tests.test_employee_capacity_utilization_report` and `node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line`.

## Authorized live verification — 2026-10-28

One 30-minute entry was created for Hassan Malik on `DIGITALLOG-1162` at 09:00 +05:00, worklog ID `69101`, with an explicit EPR integration verification description. Independent Jira reads confirmed exactly one matching entry and 1800 seconds. Jira rejected earlier attempts before creating any record because the request property needed a JSON object rather than a scalar string; that payload was corrected and regression coverage added. No further live writes are part of verification. Personal OAuth consent remains untested live until the app is registered and configured.
