# My Day

## Purpose

My Day is the personal home page of EPR. It opens after sign-in (`/` redirects to `/my_day.html`) and is pinned at the top of the sidebar. It shows one employee's day: hours logged against the hours expected, yesterday's gaps, a timesheet streak, today's subtasks with one-click logging, this week's bars, the epics they worked on this month, and upcoming leave and holidays.

## Business Logic

- **Who is shown.** With sign-in on, Members always see their own Atlassian display name. Admins, and servers with sign-in off, can switch employees with `?employee=<Name>` or the **Viewing** picker. With sign-in off and no employee chosen, the API returns `needs_employee` and the page asks for one.
- **Which day.** The current server date, or `?date=YYYY-MM-DD` (for reviewing a past day or for demoing local data).
- **Expected hours per day** (`expected_hours`): Saturday and Sunday are 0. A date in the capacity profile's `holiday_dates` is 0. A date inside the profile's Ramadan window uses `ramadan_hours_per_day`. Any other day uses `standard_hours_per_day` (default 8). The profile is the `assignee_capacity_settings` row whose `from_date`–`to_date` covers the date.
- **Day status** (`day_status`), applied in this order:
  1. after today → `future` (or `leave` when planned leave exists)
  2. 0 expected → `off`
  3. leave ≥ expected → `leave`
  4. logged + leave ≥ expected − 0.5h (`DAY_GRACE_HOURS`) → `complete`
  5. the day is today → `today`
  6. date later than the newest synced worklog (`data_through`) → `unsynced`, which is never counted as missed
  7. otherwise → `missing`, with missing hours = expected − logged − leave
- **Missed yesterday.** Walk back from yesterday, up to 10 days, to the last day with hours expected. If that day is `missing`, show it with a **Fix it now** button that pre-fills Quick log with that date and the missing duration.
- **Timesheet streak.** Look at the 26 Mon–Fri weeks before the current week, newest first. Weeks containing `unsynced` days are skipped. Count weeks in which every workday is `complete`, `leave` or `off`, and stop at the first incomplete week. The dots show the last 6 weeks (filled = complete, dashed = not synced).
- **Today's tasks.** Subtasks (`issue_type` contains "sub…task", excluding `RLT-*`) assigned to the employee are included when either of these holds:
  - Not done, the start date is on or before today (or blank), and the due date is no more than 14 days ago.
  - Hours were logged on them today.

  Done statuses are Done, Closed, Resolved, Resolved!, Bug Fixed and Not a Bug. Sort order: open before done, most logged today, overdue first, earliest due date. At most 10 tasks are shown.
- **Recent tab.** Other subtasks the employee logged on in the last 26 weeks, newest first, at most 8.
- **My epics this month.** Epics of subtasks with hours logged from the 1st of the month to today, top 3 by this month's hours. The card compares all logged hours with the original estimates on the employee's own subtasks in that epic:
  - **Over plan**: logged > planned.
  - **Just started**: < 25% of planned, or no estimate.
  - **On track**: otherwise.
- **Coming up.** Next 14 days. Shows days with RLT planned-not-taken leave first, then official holidays, at most 4 entries.
- **Live hours.** These conditions must all hold:
  - The viewed day is today.
  - The browser has a personal Jira session (`epr_jira_session`).
  - The Jira identity matches the employee.

  When they do, this week's worklogs are read live from Jira (`worklogAuthor = <accountId> AND worklogDate` between Monday and today). They replace this week's canonical hours, and the header chip reads **Live from Jira**. Otherwise the chip shows **Synced through <date>**.
- **Quick log** reuses the existing reviewed worklog flow: `POST /api/employee-capacity-utilization/worklog/review`, then a confirm checkbox, then `/submit`. It inherits all of that flow's rules: a non-leave subtask, the employee must equal the Jira identity, no future dates, a 24h cap per subtask per day, and a 1–2000 character description. After a successful submit, the ring, week, missed card and task row update in place without waiting for a refresh.

## Business Cases

Employees used EPR only when they needed a report, so there was no reason to open it daily. My Day turns EPR into the place where people start the day: they see what to work on, spot unlogged hours early, and log time faster than in Jira. Fewer missing timesheets make every downstream capacity, utilization and plan-vs-actual report more accurate. Admins measure the adoption on [Usage Analytics](../../capacity-user-guide/screens/09-usage-analytics.md).

## Examples

| Input | Result |
|---|---|
| Wed 21 Oct, 8h expected, 3h logged, data synced through 21 Oct | Missed card: "5h missing on Wed, 21 Oct" |
| Same, but data synced only through 20 Oct | Day is `unsynced`; no missed card |
| Tue 20 Oct, 7.5h logged | `complete` (within the 0.5h grace) |
| Fri 23 Oct listed in `holiday_dates` | 0h expected; shown in Coming up as "Official holiday" |
| RLT planned leave on 28 Oct, viewed on 22 Oct | Coming up: "Planned leave · 8h · RLT-9" |
| Previous week every workday complete, the week before has a gap | Streak: "1 full week in a row" |
| Click **+1h** on PLAT-482 | Quick log selects PLAT-482, duration 1h, focus moves to the description |

## Explanations

When you open EPR you land on My Day. The ring tells you how much of today's expected time you have logged. The yellow card warns you if your last workday has a gap, and one click opens a pre-filled entry for it. The purple card counts how many full weeks in a row you have logged completely. The task list shows the subtasks you are expected to be working on; the small buttons add 30 minutes or an hour, or open a custom entry. Quick log never sends anything to Jira until you review the entry and tick the confirmation. The week bars use colors: green complete, yellow missing, purple today, blue leave, dashed not yet happened or not yet synced.

## Front-end UI Fields

| Field | Type | Default | Range / values | Controls | Example |
|---|---|---|---|---|---|
| Viewing (picker) | select | signed-in user | canonical worklog authors | whose day is shown (Admins / sign-in off only) | "Faiza Nasir" |
| Freshness chip | label | — | Live from Jira / Synced through DATE | explains where hours come from | "Synced through 31 Aug" |
| Log time | button | — | — | scrolls to Quick log | — |
| Today's hours ring | SVG | — | 0–100% of expected | logged ÷ expected | "5.5h of 8h" |
| Missed card | card + button | hidden when no gap | — | Fix it now pre-fills date and duration | "2h missing on Mon, 27 Oct" |
| Timesheet streak | number + 6 dots | 0 | 0–26 weeks | — | "4 full weeks in a row" |
| Assigned to me / Recent | tabs | Assigned | — | switches task list | — |
| +30m / +1h / clock | buttons per task | — | — | pre-fill Quick log | — |
| Subtask | select | first open task | assigned, recent, done groups | worklog issue key | "PLAT-482 · Refactor sync retry" |
| Duration chips / Or type | buttons + text | 1h | 1 minute–24h; formats 30m, 1h 30m, 0.5h, 1d | worklog duration | "1h 30m" |
| Start time | time | 09:00 | 00:00–23:59 | worklog start | 14:00 |
| Date | date | viewed day | ≤ today | worklog date | 2026-10-27 |
| What did you do? | textarea | empty | 1–2000 chars, required | Jira worklog comment | "Retry backoff + tests" |
| Review entry → confirm → Log | buttons + checkbox | — | — | two-step Jira submission | — |
| This week | 5 bars | — | Mon–Fri | logged / leave per day | "19.5h / 40h" |
| My epics this month | up to 3 cards | — | on track / over plan / just started | — | "31h of 50h planned" |
| Coming up | up to 4 rows | — | next 14 days | planned leave, holidays | "THU 30 · Planned leave" |

## Script Files

| File | Role |
|---|---|
| `my_day.html` | Canonical page (HTML, CSS and inline JS); promoted to `report_html/my_day.html` by `sync_report_html()` |
| `my_day_service.py` | `build_my_day`, `expected_hours`, `day_status`, `fetch_live_week_worklogs`, `list_people` |
| `report_server.py` | `GET /api/my-day` route, `/` redirect, `_resolve_report_html_sources()` entry |
| `shared-nav.js`, `shared-nav.css` | Pinned **My Day** link (`is-pinned`) above the report list |
| `employee_worklog.py` | Reused `/identity`, `/review`, `/submit` worklog endpoints and `jira_identity()` |
| `tests/test_my_day_service.py` | Unit tests for status rules, streak, tasks, epics, live overlay |

## Dependent & Impacted Files

- `epr_auth.py`: `my_day.html` is a report file, so Members may open it; `/api/my-day` is a GET API, which Members are allowed. The `/` redirect now lands on My Day, so `safe_next('/')` after sign-in lands there too.
- `usage_analytics.py`: counts `/my_day.html` views and Quick log submissions whose Referer is My Day (`via_my_day`).
- `canonical_report_data.build_rlt_leave_snapshot`: source of leave and planned leave.
- [Employee Capacity & Utilization](16-employee-capacity-utilization.md): shares the worklog review/submit endpoints and their rules.
- `tests/test_admin_navigation.py`, `tests/test_report_html_promotion.py`: assert `/` redirects to `/my_day.html`.

## Table Schema

Read-only, `assignee_hours_capacity.db` (canonical run from `canonical_refresh_state.last_success_run_id`):

| Table | Columns used | Meaning |
|---|---|---|
| `canonical_refresh_state` | `last_success_run_id`, `updated_at_utc` | active run and refresh time |
| `canonical_issues` | `issue_key`, `summary`, `status`, `issue_type`, `assignee`, `start_date`, `due_date`, `original_estimate_hours`, `total_hours_logged`, `epic_key` | subtasks and epic names |
| `canonical_worklogs` | `issue_key`, `worklog_author`, `started_date`, `hours_logged` | hours per subtask per day |
| `assignee_capacity_settings` | `from_date`, `to_date`, `standard_hours_per_day`, `ramadan_*`, `holiday_dates_json` | expected hours calendar |
| RLT canonical rows (via `build_rlt_leave_snapshot`) | daily `planned_taken_hours`, `unplanned_taken_hours`, `unknown_taken_hours`, `planned_not_taken_hours` | leave taken and planned |

No table is written. Quick log writes only to Jira (and the existing worklog receipt files).

## Data Flow

1. The browser loads `/my_day.html`; `shared-nav.js` checks `/api/auth/me`.
2. The page calls `GET /api/my-day[?employee=&date=]`. `report_server.my_day_api` resolves the employee from `g.epr_user` (or the parameter for Admins and sign-in-off servers).
3. If live conditions hold, `fetch_live_week_worklogs` calls Jira `search/jql` plus `issue/<key>/worklog` for this week.
4. `build_my_day` reads the canonical tables and the leave snapshot, applies the rules above, and returns JSON. The route adds `can_switch`, `people` and `jira_browse_base`.
5. The page renders the cards and calls `/worklog/identity` to enable Quick log.
6. Quick log → `/worklog/review` (signed token) → confirm → `/worklog/submit` → Jira worklog. The page updates its numbers locally, and `usage_analytics` records `worklog_submitted`.

## Local verification

```powershell
.venv/Scripts/python.exe -m unittest tests.test_my_day_service tests.test_usage_analytics
.venv/Scripts/python.exe run_server.py
```

Open `http://127.0.0.1:3000/my_day.html?employee=Faiza%20Nasir&date=2026-08-27`; local canonical data ends at the last refresh, so pass a `date` inside it.

## Change notes

- 2026-10-28: first release (My Day home page, pinned nav link, `/` redirect).
