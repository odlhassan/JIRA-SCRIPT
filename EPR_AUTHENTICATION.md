# EPR authentication and access

EPR reuses the existing Atlassian OAuth integration for application sign-in and personal Jira worklogs. No passwords, additional OAuth scopes, package dependencies, or SQLite schema changes were introduced. This is the implemented two-role core; the older password-based AUTH_ROLE design drafts do not describe current behavior.

## User experience and roles

Signed-out report requests redirect to `/login`; APIs return 401 without report data. **Continue with Atlassian** uses the registered Jira site and browser-bound OAuth state check, then returns to the original local report URL. First-time users reach **Access needs approval**. Signing in alone grants no report access.

An Admin approves the user in **Admin Settings → Users & access**. Members read reports and submit personal worklogs through the existing review and confirmation flow. Admins manage settings, refreshes, plans, SQL, and access. Approved Members see organization-wide report data; this core release does not add team/row-level confidentiality.

Hassan Malik's verified Atlassian ID is the bootstrap Admin. `EPR_ADMIN_ACCOUNT_IDS` can explicitly replace that comma-separated bootstrap list; it only applies when creating a missing access record. Users must sign in once before approval. Role, Approved, and Active changes apply on the next request, including existing sessions. The server prevents disabling your own account and removing the last approved, active Admin.

### Return destination after sign-in

`/login?next=` stores the requested page in the short-lived `epr_next` cookie, and the OAuth callback redirects there. `safe_next()` accepts only navigable EPR pages, such as `/`, `*.html`, `/settings/...`, `/my-availability` and `/report_html/`. It rejects the following, and each falls back to `/`:

- external or protocol-relative URLs, backslashes and control characters
- `/api/...` paths
- `/login`, `/access-pending`, `/healthz` and `/favicon.ico`
- `/fonts/...` and `/project-images/...` paths
- any file with an extension other than `.html`

`/login` writes the cookie only on a top-level document load (`Sec-Fetch-Dest: document`, or the header missing). It also skips the write when the request carries a rejected `next`. Background requests, such as the browser's automatic favicon fetch, therefore cannot overwrite a deep link. `/favicon.ico` is public and returns 204, so it never reaches sign-in. This fixed production sign-ins that ended on `/favicon.ico` with `{"error":"Not found"}`. A stale asset value still in a cookie resolves to `/`.

Navigation displays name, role, **My availability**, and **Sign out** in the sidebar footer (see `SHARED_NAVIGATION.md`). My availability opens that employee's calendar in the report's current month. Jira remains the author of record; another employee's calendar cannot be used to submit on their behalf.

The light login experience uses a glass surface, liquid-gradient orb, drifting pastel nebulae, sparse stars, and an illustrative network of task/milestone/chart cards. Graphics are decorative, with no real project metrics or private data before login. Pointer movement creates subtle parallax and nearby star repulsion; clicks create soft ripples. The motion control and system reduced-motion preference pause the ambience. Canvas is capped at 30fps, 65 stars (30 on mobile), and pixel ratio 1.5; it stops when the tab is hidden. No external graphics packages, fonts, video, or image services are loaded. `epr-login.css` and `epr-login.js` are public exact-match assets served by `epr_auth.py`; other protected files stay behind the gate.

## Server policy

`epr_auth.py` protects direct URLs and APIs, independently of menu visibility. Members have report GET/HEAD access except administrative routes. Mutations default to denied, with explicit exceptions for personal worklog review/submit, the two report exports, and capacity calculation. Settings pages and access/audit APIs require Admin. New mutations need an intentional policy update. Some existing report controls may remain visible; unauthorized requests receive 403.

The existing opaque HttpOnly, SameSite=Lax cookie is Secure on HTTPS. EPR sessions expire eight hours after identity validation. Role/approval are read from private storage on each request. Reauthentication creates a new EPR session. Mutations require a same-origin Origin header or the server-issued `X-EPR-CSRF` token as well as authentication. Authenticated responses are private/no-store with nosniff and same-origin framing; wildcard CORS is disabled.

## Private storage and privacy

State is beside the capacity DB under `worklog_requests`, normally `/home/data/worklog_requests`:

- `oauth/<random-session>.json`: OAuth tokens, verified cached Jira identity, EPR sign-in time and CSRF token.
- `epr_access/users.json`: account ID/name, role, approved/active state and retrieval timestamp.
- `epr_access/audit.jsonl`: login/logout, access requests/denials, access changes and authenticated mutations. Bodies, query strings, tokens and secrets are excluded. The screen shows 200 recent events; the persisted file currently has no rotation.

JSON files are atomically replaced with private permissions where supported. Keep the existing single Gunicorn worker. Back up this directory securely; never serve or commit it. Logout removes integration sessions but preserves access approvals. Privacy `closed`/`updated` instructions remove profile/access records and anonymize matching audit IDs, requiring fresh approval. Reporting uses an existing valid OAuth grant; when none exists it waits for an authorized session. Jira worklogs and canonical reporting data are unchanged by integration cleanup.

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
