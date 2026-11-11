# EPR authentication and access

EPR reuses the existing Atlassian OAuth integration for application sign-in and personal Jira worklogs. No passwords, additional OAuth scopes, package dependencies, or SQLite schema changes were introduced. This is the implemented two-role core; the older password-based AUTH_ROLE design drafts do not describe current behavior.

## User experience and roles

Approved Members and Admins may delete **their own** Jira worklog entries through `/api/employee-capacity-utilization/worklog/delete`. `employee_worklog.py` requires a valid personal OAuth session and rechecks the live worklog author account ID; Admin status grants no override. The manager in My Day and Availability uses this specific Member POST allowance with existing same-origin controls. Read/list operations use the same personal identity; shared server credentials cannot manage entries. See [worklog manager data flow](JIRA_WORKLOG_SIGNIN.md#delete-your-own-logged-hours).

Signed-out report requests redirect to `/login`; APIs return 401 without report data. **Continue with Atlassian** uses the registered Jira site and browser-bound OAuth state check, then returns to the original local report URL. First-time users reach **Access needs approval**. Signing in alone grants no report access.

An Admin approves the user in **Admin Settings → Users & access**. Members read reports and submit personal worklogs through the existing review and confirmation flow. Admins manage settings, refreshes, plans, SQL, and access. Approved Members see organization-wide report data; this core release does not add team/row-level confidentiality.

Hassan Malik's verified Atlassian ID is the bootstrap Admin. `EPR_ADMIN_ACCOUNT_IDS` can explicitly replace that comma-separated bootstrap list; it only applies when creating a missing access record. Users must sign in once before approval. Role, Approved, and Active changes save automatically when changed and apply on the next request, including existing sessions. Approval and active-state changes require confirmation. The management dialog shows saving, saved, or an error; failed changes revert to their previously saved values. The server prevents disabling your own account and removing the last approved, active Admin.

### Users & access directory

The People tab shows counts of all people, pending approvals, and active access. Search matches name or Atlassian account ID; filters narrow by access status and role. Suspended means inactive, pending means active but unapproved, and active means both approved and active. The status is descriptive and does not change the underlying approval and active flags. Select Review access or Manage access to open the person's role and access controls. The dialog shows the account ID and offers a copy action. Changes save automatically; approval and active changes ask for confirmation, and rejected changes revert to their saved values. The Activity tab displays the latest 200 audit events and expands technical details. Successful changes refresh activity. User data comes from `/api/auth/users`; changes use `PUT /api/auth/users/<account_id>`, and activity comes from `/api/auth/audit`. The page is `epr_users.html`; authorization and validation remain in `epr_auth.py`. Access records and audit events are private files, not SQLite tables.

### Return destination after sign-in

Without a `next`, sign-in returns to `/`, which redirects to **My Day** (`/my_day.html`, a report file Members may open; see `docs/report-user-guide/screens/17-my-day.md`). `/login?next=` stores the requested page in the short-lived `epr_next` cookie, and the OAuth callback redirects there. `safe_next()` accepts only navigable EPR pages, such as `/`, `*.html`, `/settings/...`, `/my-availability` and `/report_html/`. It rejects the following, and each falls back to `/`:

- external or protocol-relative URLs, backslashes and control characters
- `/api/...` paths
- `/login`, `/access-pending`, `/healthz` and `/favicon.ico`
- `/fonts/...` and `/project-images/...` paths
- any file with an extension other than `.html`

`/login` writes the cookie only on a top-level document load (`Sec-Fetch-Dest: document`, or the header missing). It also skips the write when the request carries a rejected `next`. Background requests, such as the browser's automatic favicon fetch, therefore cannot overwrite a deep link. `/favicon.ico` is public and returns 204, so it never reaches sign-in. This fixed production sign-ins that ended on `/favicon.ico` with `{"error":"Not found"}`. A stale asset value still in a cookie resolves to `/`.

Navigation displays name, role, **My Day**, **My availability**, and **Sign out** in the sidebar footer (see `SHARED_NAVIGATION.md`). My availability opens that employee's calendar in the report's current month. Jira remains the author of record; another employee's calendar cannot be used to submit on their behalf.

The light login experience uses a glass surface, liquid-gradient orb, drifting pastel nebulae, sparse stars, and an illustrative network of task/milestone/chart cards. Graphics are decorative, with no real project metrics or private data before login. Pointer movement creates subtle parallax and nearby star repulsion; clicks create soft ripples. The motion control and system reduced-motion preference pause the ambience. Canvas is capped at 30fps, 65 stars (30 on mobile), and pixel ratio 1.5; it stops when the tab is hidden. No external graphics packages, fonts, video, or image services are loaded. `epr-login.css` and `epr-login.js` are public exact-match assets served by `epr_auth.py`; other protected files stay behind the gate.

## Server policy

Deployment maintenance sits outside this authentication layer. While paused, only
the maintenance shell/assets, game/status APIs and liveness route reach their
database-free handlers; report/auth mutations receive 503. When reopened, existing
authentication and authorization apply normally. The game remains intentionally
anonymous during healthy operation, reachable through a privately shared shortcut.
There is no game link/hint in ordinary navigation.

The bootstrap owner (`OWNER_ACCOUNT_ID`) has an `is_superadmin` capability only while
approved, active and Admin. Only this account can manage `/api/auth/game-shortcut`
and see the preference panel in Users & access. Other Admins and Members receive
403 from that endpoint. `/api/game-shortcut` is readable by approved app users;
normal same-origin/CSRF checks protect saves. The preference is an atomically written
private JSON file, with no new role or SQLite schema. See
[maintenance and private playground](MAINTENANCE_GAME.md).

### Dependent & Impacted Files — private shortcut

`epr_users.html` loads `epr-superadmin-preferences.js`, which creates its panel only
after verified identity reports the owner capability. `epr_auth.py` serves
`app-shortcuts.js`; shared navigation and Seating Planner load that common listener.
`maintenance_service.py` handles the anonymous room outside the reporting auth gate.
The game serves the same listener at `/maintenance/app-shortcuts.js`. Its public
status exposes only the validated, non-secret key via the auth module's read-only
shortcut callback, so anonymous visitors can use the configured return shortcut.
Preference writes and application access remain under the existing owner/auth checks.

### Data Flow — private shortcut

Owner identity → owner-only preference endpoint → private
`epr_access/game_shortcut.json` → approved-user shortcut endpoint → hidden key listener
→ anonymous playground. Returning to reports repeats ordinary access checks.
The management API checks role/approval/active status and owner account independently
of panel visibility; browser visibility does not grant permission.
On return, the listener checks live availability and restores the local path,
query and anchor. Report authentication still runs normally. The client caches
only the last known shortcut in sessionStorage, alongside its anonymous nickname.

`epr_auth.py` protects direct URLs and APIs, independently of menu visibility. Members have report GET/HEAD access except administrative routes. Mutations default to denied, with explicit exceptions for personal worklog review/submit, the two report exports, and capacity calculation. Settings pages and access/audit APIs require Admin, including **Usage Analytics** (`/settings/usage-analytics` and `/api/admin/usage-analytics`, via the `/admin/` rule). `GET /api/my-day` is a normal Member API; a Member always gets their own day and the `employee` parameter is honored only for Admins. New mutations need an intentional policy update. Some existing report controls may remain visible; unauthorized requests receive 403.

The existing opaque HttpOnly, SameSite=Lax cookie is Secure on HTTPS. EPR sessions last up to 30 days and never outlive the Jira session's existing 30-day absolute expiry. Existing sessions less than 30 days old can continue without another eight-hour login. Role/approval are read from private storage on each request. Jira access tokens refresh server-side; refreshes are serialized within the production single worker so concurrent requests do not race rotating credentials. Expiry or an invalid grant removes the affected session rather than other signed-in sessions. Sign-out, disabled access, revoked grants, privacy erasure, browser cookie removal, and the 30-day expiry can still require a fresh sign-in. Atlassian requires the consent prompt when a new authorization flow is started. Mutations require a same-origin Origin header or the server-issued `X-EPR-CSRF` token as well as authentication. Authenticated responses are private/no-store with nosniff and same-origin framing; wildcard CORS is disabled.

## Private storage and privacy

State is beside the capacity DB under `worklog_requests`, normally `/home/data/worklog_requests`:

- `oauth/<random-session>.json`: OAuth tokens, verified cached Jira identity, EPR sign-in time and CSRF token.
- `epr_access/users.json`: account ID/name, role, approved/active state and retrieval timestamp.
- `../usage_analytics.db` (beside the capacity DB): page views and EPR worklog submits per signed-in user (account ID, name, role, date, path), pruned after 400 days. It is used only by the Admin Usage Analytics screen. See `docs/capacity-user-guide/screens/09-usage-analytics.md`.
- `epr_access/audit.jsonl`: login/logout, access requests/denials, access changes and authenticated mutations. Bodies, query strings, tokens and secrets are excluded. The screen shows 200 recent events; the persisted file currently has no rotation.

JSON files are atomically replaced with private permissions where supported. Keep the existing single Gunicorn worker. Back up this directory securely; never serve or commit it. Logout, expiry, revoked tokens, and reauthentication preserve the stored approval and role. An administrator can still revoke access. A privacy `updated` instruction removes stale profile data and OAuth sessions while retaining the approval decision tied to the same account ID; the next sign-in restores the current profile. A `closed` instruction removes access and anonymizes matching audit IDs. Records already removed before this change cannot be reconstructed from the session. Reporting uses an existing valid OAuth grant; when none exists it waits for an authorized session. Jira worklogs and canonical reporting data are unchanged by integration cleanup.

## Configuration and operation

Authentication defaults on in Azure (`WEBSITE_SITE_NAME`) or when all three existing `JIRA_OAUTH_*` settings exist. `EPR_AUTH_ENABLED=1` enables local gate verification without secrets. An unconfigured local server otherwise retains its previous behavior; Azure has no such bypass. Public `/healthz` returns minimal liveness JSON and should be used for health probes. Public `/favicon.ico` returns 204.

The existing production OAuth callback, client configuration, scopes, and internal distribution remain valid. Real localhost OAuth needs a separately registered matching local callback; the production HTTPS callback cannot authenticate localhost. No database migration is pending.

Implementation spans `epr_auth.py`, `jira_oauth.py`, `jira_privacy.py`, `report_server.py`, the three `epr_*.html` screens, and `shared-nav.js`/`.css`. The employee report generator adds the personal-calendar entry URL; regenerate its root and promoted copies together.

## Local verification

```powershell
.venv/Scripts/python.exe -m unittest tests.test_epr_auth tests.test_employee_worklog tests.test_jira_privacy tests.test_employee_capacity_utilization_report
node node_modules/@playwright/test/cli.js test --config tests/epr-auth.playwright.config.cjs --workers=1 --reporter=line
node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line
node --check shared-nav.js
node --check epr-login.js
.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py
Copy-Item -LiteralPath employee_capacity_utilization_report.html -Destination report_html/employee_capacity_utilization_report.html
$env:EPR_AUTH_ENABLED='1'
.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3121
```

Open `http://127.0.0.1:3121/employee_capacity_utilization_report.html` and expect login. Without local OAuth configuration, Continue is disabled with an explanation. API/browser tests use temporary identities and mock Jira writes. Production checks must verify anonymous denial and Hassan's existing Admin session without submitting another worklog.
