# Deployment maintenance and private playground

## Business Cases

Pause the app before replacing its code, finish active requests and report jobs, and
reopen after the new revision initializes. Visitors can play anonymous Chorr Police
while waiting. App users can also enter through a privately shared shortcut during
healthy operation. This uses the existing App Service instance; no new hosting
service, dependency, game database or schema migration is introduced.

## Business Logic

The outer WSGI gate runs before authentication. During maintenance, document requests
receive the playground with HTTP 503 and Retry-After: 5; application APIs, assets and
mutations receive 503 without entering the core app. Exact public routes are
`/maintenance`, its four assets, `/api/maintenance/status`, `/api/maintenance/game`
and `/healthz`. Health is a liveness probe; deployment readiness requires
`status.ready=true` and the matching packaged `DEPLOYMENT_REVISION`.

State lives in an atomically replaced JSON marker:
`/home/data/deployment-maintenance.json` on Azure, `.deployment-maintenance.json`
at the project root locally. `EPR_MAINTENANCE_FILE` overrides the local path. SCM
uses the fixed production path, so retain that default on Azure. No marker means
available once startup succeeds; invalid/unreadable markers and startup failure
keep the app paused. Phases are available, draining, deploying, checking and failed.
Maintenance never times out to available automatically.

One Gunicorn worker with eight threads serves the gate while the core initializes
in a startup thread. The gate counts responses until their iterators close.
`start_background_job` counts jobs before their HTTP response finishes and decrements
even on failure. This covers report, canonical Fetch/Compute/assignee/epic, missed
entries, performance, planned-actual, dashboard and offline-preparation jobs.
Existing periodic OAuth/privacy services are outside this job count.

One locked in-memory room holds up to 40 players. The first entrant is police;
later entrants are runners. The server advances bounded directional input using
elapsed monotonic time in collision steps of at most 1/60 second, normalizes
diagonals and blocks houses/jail/map edges. A 200 ms gap covers the full 200 ms
of travel; speed no longer depends on participant count or polling frequency.
Police catch runners within 34 map units. Caught runners appear in jail. A free
runner touching its boundary frees prisoners with three seconds of immunity.
Catching every runner starts another round after five seconds. Free runners may
request a badge; either officer can approve/decline. Acceptance is atomic, with
two police maximum. Police cannot be caught.

Movement stops after 0.6 seconds without input. Players expire after 20 seconds
without polling; if no police remain, the oldest participant takes the badge and
leaves jail if necessary. Rejoining with an existing cookie is idempotent. Room
state resets on worker/container restart; connected browsers rejoin automatically.
No Jira identities or report data enter the game. Commands require the site's
Origin and JSON, limited to 2 KiB; private opaque cookie tokens never appear in
the room snapshot. Anonymous access is intentional.

Drawing is independent of the network heartbeat. `motion.js` predicts the local
player on each display frame, immediately changes direction on key/pointer edges,
and reconciles small server differences over 120 ms instead of jumping. Other
players interpolate a short snapshot history with a buffer of 200 ms plus measured
round-trip latency (capped at a 400 ms buffer), with at most 100 ms extrapolation.
Jail, rescue, role, room and round transitions use the authoritative position
immediately. Prediction stops after 600 ms without a response. Server validation,
catches, team approval and jail releases remain authoritative.

The browser uses its display animation rate (typically 60 or 120 Hz), caches the
static neighbourhood drawing and sizes player rendering for device pixel density
up to 2. Idle network heartbeats remain 200 ms while joined and one second while
outside the game. Input edges send immediately, with at most one movement request
in flight. Blur/tab hiding sends a zero direction. No extra service or animation
dependency is needed.

`Ctrl+G` toggles between the app and the same room while healthy, preserving the
original app path, query filters and anchor. A second press in the game first
checks fresh availability; a paused app or failed check keeps the game open.
Successful return stops polling and sends a keepalive leave command. The key
works from the game's nickname field too; ordinary app typing controls retain
their existing behavior. Held-key repeats cannot toggle repeatedly. There
is no menu item, tooltip, help hint or displayed combination for ordinary users.
The listener ignores typing/editable controls and repeated/composing key events.
Preferences refresh every minute, on tab visibility and after an owner saves.
The game receives the non-secret combination through public availability status,
so anonymous visitors also get the configured key without a report login. The
last known combination is cached in sessionStorage across navigation/restarts;
the game's normal status polls refresh it too. Management remains owner-only.
Without a cached/configured value during startup, the waiting page uses Ctrl+G
until the initialized app supplies its configured combination.
The known direct URL also works: undisclosed access is a presentation preference,
not an authorization boundary.

Only the approved, active Admin matching `OWNER_ACCOUNT_ID` has the superadmin
capability for this preference. This does not introduce another persisted role.
Other Admins and Members cannot manage it. Approved users read `/api/game-shortcut`;
the owner reads/writes `/api/auth/game-shortcut` using normal origin/CSRF protections.
Valid combinations contain Ctrl or Alt, optional Shift, and one letter/digit.
Some operating-system/browser-reserved combinations may not be interceptable.

## Examples

- Officer joins before Samosa and Mango: Officer is police, both others are runners.
- Officer catches Samosa; Mango touches jail, freeing Samosa with temporary immunity.
- Mango requests police membership; Officer approves. A third badge is rejected.
- The owner saves Ctrl+Shift+H. Open app pages adopt it within a minute without hints.
- Deployment installs revision abc123. A ready old revision cannot reopen the app.

## Explanations

People see a thin full-width availability notice above the game. They enter a
nickname and play without sign-in. The notice follows actual deployment phases,
without an invented percentage or completion time. Once deployment finishes, waiting
users return to their original app page and normal authentication applies. A visit
made while already healthy stays in the game until Return to app is selected or
the same privately shared shortcut is pressed again.

## Front-end UI Fields

- Availability strip: current status, checked every five seconds; failure says the
  update needs attention. Return to app appears when available and uses a safe local URL.
- Nickname: required by the form, maximum 18 characters; API default Guest. Rendered
  as text, never HTML. Browser sessionStorage holds the nickname and last known
  non-secret shortcut, without Jira identity or authentication credentials.
- Enter playground: joins once; hides the welcome overlay; admission stops at 40.
- Online count, round, role, line-up: server state; your map player has a ring and
  jailed players show IN JAIL.
- WASD/arrows and direction buttons: movement; release, blur and tab hiding clear
  input. Small screens show direction buttons; coarse pointers also receive them.
  Joining focuses the map so movement works immediately after Enter playground.
- Ask to join police: free runners only; disabled while awaiting approval; hidden
  at two officers. Accept/Decline appear only for police. Leave removes the participant.
- Private playground shortcut: owner-only panel under Users & access. Keyboard
  shortcut is required text, maximum 32 characters. Save shortcut validates,
  persists and shows Saving/Saved/error. Rejection preserves the saved setting.

## Script Files

- `maintenance_service.py`: WSGI gate, endpoints, room, job tracker, marker CLI.
- `maintenance/index.html`, `game.css`, `game.js`: shell, map, controls, polling/reconnect.
- `maintenance/motion.js`: collision-safe local prediction, server reconciliation
  and smooth remote interpolation; loaded before game.js and served without auth.
- `azure_deploy_maintenance.py`: SCM marker and drain/revision handshake.
- `.github/workflows/azure-appservice-deploy.yml`: packaging and maintenance lifecycle.
- `wsgi.py`, `startup.txt`: core warmup behind the gate in one threaded worker.
- `app-shortcuts.js`: shared hidden toggle listener and 10-second maintenance
  handoff. The game loads the same file via `/maintenance/app-shortcuts.js`.
- `epr-superadmin-preferences.js`, `epr_users.html`: owner settings UI.
- `tools/preview_maintenance.py`: database-free preview on port 3096.
- `tests/test_maintenance_service.py`, `test_maintenance_shortcut.py`,
  `test_azure_deploy_maintenance.py`, `test_maintenance_browser.cjs`: focused checks.
- `tests/test_maintenance_motion.cjs`: frame-rate independence, immediate input,
  remote interpolation, collision, jail/round transitions and stale-response tests.

## Dependent & Impacted Files

`report_server.py` wraps the local app and tracks report jobs. `shared-nav.js`
loads `app-shortcuts.js`; Seating Planner explicitly loads it because it omits the
sidebar. `epr_auth.py` serves the assets/preferences and enforces owner access.
See `AZURE_APP_SERVICE.md`, `SHARED_NAVIGATION.md`, `EPR_AUTHENTICATION.md`.
These assets are canonical root files. Restart updates server behavior; no report
regeneration is needed, and ordinary startup sync updates shared-nav.js.
`epr_auth.py` supplies a read-only shortcut callback through the app's extensions;
the local gateway receives it explicitly and the warmed production app exposes it
to its outer gateway. Only the validated combination reaches public status.

## Table Schema

No SQLite tables are added, changed, read or written by this module. Runtime JSON
files are the deployment marker and private
`worklog_requests/epr_access/game_shortcut.json` containing only `{"shortcut":"Ctrl+G"}`.
Game memory holds public player IDs/names/roles/positions/jail/request flags and
private opaque tokens. Existing private-file permissions protect the preference.

## Data Flow

1. GitHub stages code, writes GITHUB_SHA into DEPLOYMENT_REVISION, vendors dependencies,
   verifies assets and builds the ZIP before enabling maintenance.
2. SCM writes a run-owned draining marker. New core requests stop.
3. The workflow requires two consecutive idle polls: active_requests=0 and
   background_jobs=0, ready=true. A 20-minute timeout aborts and leaves maintenance on.
4. The workflow marks deploying, uploads the ZIP, then marks checking.
5. The new WSGI gate serves the public shell while initializing the core app.
6. Only ready=true, the new revision and matching deployment owner allow available.
7. Waiting browsers restore their original destination. Failure/cancellation stays paused.
8. Independently, users load the saved hidden shortcut and enter the same room.
   Cookie-bound directions/actions update the authoritative room under a lock.
   Snapshots include room ID, monotonic simulation time, pause state, public movement
   directions and the collision/speed rules. The browser renders predicted/interpolated
   positions between those snapshots; it never submits absolute player coordinates.
9. In the room, the same shortcut reads fresh availability, stops movement/polling,
   sends leave and restores the safe original URL. Public status supplies only
   the validated shortcut; it grants no report access or settings write permission.

## First rollout, limitations and recovery

The previous release cannot show a gate it does not contain. For the first installation,
manually dispatch **Azure App Service deploy** with **bootstrap_maintenance** enabled.
This explicitly permits the old release's missing endpoint. Later pushes use the
normal guarded path. Without bootstrap, a missing gate aborts deployment.

This supports one App Service instance and one Gunicorn worker. Scale-out requires
shared game/drain coordination. CPU and network usage consume the current plan;
no separate hosting service is provisioned. An Azure container restart can briefly
interrupt connections and reset the room. Loaded pages reconnect; new visitors may
temporarily receive Azure's response. Continuous service through replacement needs
separate hosting or an appropriate slot design; see
[Microsoft's ZIP package restart behavior](https://learn.microsoft.com/en-us/azure/app-service/deploy-run-package).

The base compute charge on a fixed dedicated plan is based on its configured
instances, rather than game animation frames. Extra internet egress or scale-out
can still raise the bill; the actual Azure billing/autoscale configuration has not
been inspected. See [App Service plan billing](https://learn.microsoft.com/en-us/azure/app-service/overview-hosting-plans)
and [bandwidth pricing](https://azure.microsoft.com/en-us/pricing/details/bandwidth/).

For failure/cancellation, inspect Actions/Azure logs, fix the cause and retry the
deployment. Do not remove a production marker while code is unhealthy or work is
active. There is no public toggle. SCM checks deployment ownership before reopening.

## Local use and verification

Use the repository `.venv`; bare python currently lacks Flask on this machine.
From `E:/JIRA SCRIPT`, start a preview without loading Jira or report databases:

```powershell
.venv/Scripts/python.exe tools/preview_maintenance.py --port 3096
```

Open http://127.0.0.1:3096/maintenance in two browser profiles/private sessions.
Enter different nicknames. First is police, second runner. Move, request and approve
a badge. Use three sessions with two runners for jail/rescue. The preview deliberately
does not load the reporting app or owner settings.

In another terminal, exercise maintenance and return-to-app:

```powershell
.venv/Scripts/python.exe maintenance_service.py deploying
.venv/Scripts/python.exe maintenance_service.py available
```

The notice updates on its next poll; reopening returns waiting players to `/`.
The local marker is ignored by Git. Use `--file`/`EPR_MAINTENANCE_FILE` for isolation.
To run the actual app with current reports:

```powershell
.venv/Scripts/python.exe run_server.py --keep-cache --no-sync --port 3098
```

Open http://127.0.0.1:3098/introduction.html and press Ctrl+G outside text fields.
Press Ctrl+G again to restore that report with its filters and anchor; Return to
app does the same. During maintenance, the key keeps you in the game. Change the key under
http://127.0.0.1:3098/settings/users as the signed-in bootstrap owner, with existing
Atlassian OAuth configured. Normal production sign-in works after rollout.

Focused checks executed during implementation:

```powershell
.venv/Scripts/python.exe -m unittest discover -s tests -p 'test_*maintenance*.py' -q
node tests/test_maintenance_browser.cjs
node --test tests/test_maintenance_motion.cjs
node --check maintenance/game.js
node --check maintenance/motion.js
node --check app-shortcuts.js
node --check shared-nav.js
node --check epr-superadmin-preferences.js
.venv/Scripts/python.exe -m unittest discover -s tests -p test_azure_deploy_contract.py -v
.venv/Scripts/python.exe -m unittest discover -s tests -p test_run_server.py -v
.venv/Scripts/python.exe -m unittest discover -s tests -p test_epr_auth.py -q
.venv/Scripts/python.exe -m unittest discover -s tests -p test_canonical_storage.py -q
.venv/Scripts/python.exe -m unittest discover -s tests -p test_canonical_refresh_api.py -q
node node_modules/@playwright/test/cli.js test --config tests/epr-auth.playwright.config.cjs --workers=1 --reporter=line
git diff --check
```

The browser check starts/stops its own preview on port 3097 with an isolated temporary
marker. It checks multiplayer, touch movement, approval, healthy hidden access,
owner-only preferences, maintenance handoff and automatic return. Set
MAINTENANCE_TEST_PORT if 3097 is busy. Screenshots land in tmp/maintenance-desktop.png
and tmp/maintenance-mobile.png. Local checks perform no SCM writes or real deployment.

The browser regression also records actual painted player positions. With replies
delayed by 120 ms, local movement must start within 100 ms and advance on at least
10 frames during a 460 ms hold, with no packet-sized jumps. The remote player must
also advance between network packets. Python checks compare the same travel at
1, 2, 12 and 40 polls and prove stale input cannot tunnel through a house.
It also verifies default/custom shortcut round trips, query/anchor preservation,
leave-on-return, anonymous custom keys, nickname-field return and refusal to exit
when deployment status is unavailable or cannot be checked.
