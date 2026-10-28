# Usage Analytics (Admin)

## Purpose

`/settings/usage-analytics` shows Admins how EPR is being adopted: daily and weekly active people, stickiness, My Day reach, worklogs submitted through EPR, the daily trend, the most visited pages, adoption per performance team, and approved users who have gone quiet. It is listed under **Admin Settings → Usage Analytics** in the sidebar.

## Business Logic

- **What is recorded** (`usage_analytics.track_usage`, a Flask `after_request` hook):
  - `page_view`: a `GET` that returns `200 text/html` for a top-level document load (`Sec-Fetch-Dest: document` or the header is missing). It excludes `/api/*`, `/`, `/login` and `/access-pending`.
  - `worklog_submitted`: a `POST /api/employee-capacity-utilization/worklog/submit` that returns 200 with a `worklog_id`. The stored path is the Referer page, for example `/my_day.html`.
  - With sign-in on, only approved signed-in users are recorded (`g.epr_user`). With sign-in off (local), events are stored as "Local (no sign-in)".
  - Paths are normalized: query strings are dropped, `/report_html/x.html` becomes `/x.html`, and trailing slashes are removed.
- **Person key**: the Atlassian `account_id`, or the display name when there is no account.
- **KPIs** (range = 7, 30 or 90 days ending today; "previous" = the same length immediately before):
  - **Daily active**: people with a page view today. The delta compares the average per workday over the last 7 days with the 7 days before.
  - **Weekly active**: people active in the last 7 days, compared with the 7 days before.
  - **Stickiness**: average daily active over the last 7 days' workdays ÷ weekly active, as a %.
  - **My Day reach**: people who opened `/my_day.html` in the range ÷ people active in the range.
  - **Worklogs via EPR**: `worklog_submitted` events in the range; `via_my_day` counts those whose path is `/my_day.html`.
- **Daily chart**: distinct people per day. Weekend bars are light. The first day `/my_day.html` was ever viewed is coral (the launch marker).
- **Most visited pages**: top 8 paths by views in the range, with distinct people. Titles come from `/api/page-categories` `page_catalog`, falling back to the path.
- **Adoption by team**: for each `performance_teams` team with members, the share of members (matched by display name, case-insensitive) active on at least 3 distinct days in the last 7 days.
- **Needs a nudge**: approved, active users from `epr_access/users.json` whose last page view is 5 or more Mon–Fri workdays ago, or who never opened EPR. Never-seen users are listed first, then the longest idle; at most 12 are shown, with the total count.
- **Retention**: rows older than 400 days are deleted when the server starts.
- **Export CSV**: done in the browser, with daily active users and top pages for the selected range.

## Business Cases

Stickiness work (My Day, nudges, faster logging) needs evidence. This screen tells Admins whether people come back daily, whether My Day is the entry point, which reports are actually used, and which teams or people need a reminder. It is for EPR owners and team leads in weekly reviews. It deliberately has no public ranking of individuals.

## Examples

| Data | Result |
|---|---|
| Ayesha viewed My Day on 26, 27 and 28 Oct; Bilal viewed Dashboard on 12 Oct | 30-day active people = 2; My Day reach = 50% |
| One Quick log submit from My Day on 28 Oct | Worklogs via EPR = 1 (1 from My Day) |
| Platform team = Ayesha and Bilal; only Ayesha active on 3 of the last 7 days | Platform 50% · 1/2 |
| Bilal last seen Mon 12 Oct, today Wed 28 Oct | Needs a nudge: "Last seen 12 workdays ago" |
| An approved user with no events | "Never opened EPR", listed first |

## Explanations

Every time someone opens an EPR page, the server notes who opened which page on which day. Nothing about the page content, filters or query string is stored. This screen adds those notes up. The top tiles answer "how many people use EPR daily and weekly, and do they come back?" The bar chart shows the trend and marks the day My Day went live. The two lists show which pages matter and which teams have adopted EPR. The bottom list helps Admins decide whom to remind.

## Front-end UI Fields

| Field | Type | Default | Range | Controls | Example |
|---|---|---|---|---|---|
| Range (7d / 30d / 90d) | toggle | 30d | 7, 30, 90 | `days` query parameter | 90d |
| Export CSV | button | — | — | downloads `epr_usage_<from>_<to>.csv` | — |
| Daily active / Weekly active / Stickiness / My Day reach / Worklogs via EPR | KPI tiles | — | counts and % | delta chip vs previous period | "69% · +11 pts vs before" |
| Daily active users | bar chart | — | one bar per day | hover title shows the count | "13 Oct: 27 active" |
| Most visited pages | bar list | — | top 8 | — | "My Day · 1,240 views · 41 people" |
| Adoption by team | bar list | — | 0–100% | — | "QA 64% · 4/6" |
| Needs a nudge | person cards | — | up to 12 | — | "Last seen 9 workdays ago" |

## Script Files

| File | Role |
|---|---|
| `usage_analytics.py` | Schema, `record_event`, `build_usage_summary`, `register_usage_analytics` (tracking hook, page route, API) |
| `usage_analytics.html` | Admin page (HTML, CSS and inline JS), served from the repo root by `/settings/usage-analytics` |
| `report_server.py` | Calls `register_usage_analytics` after `register_epr_auth`, passing the `performance_teams` loader |
| `shared-nav.js` | Adds the **Usage Analytics** item to Admin Settings |
| `tests/test_usage_analytics.py` | Summary math, tracking hook, admin-only access |

## Dependent & Impacted Files

- `epr_auth.py`: provides `g.epr_user`, `EPR_AUTH_ENABLED` and `EPR_AUTH_DIRECTORY/users.json`. `/settings/*` and paths containing `/admin/` are Admin-only, so Members get 403 on both the page and the API.
- [My Day](../../report-user-guide/screens/17-my-day.md): its page views and Quick log submits feed the My Day KPIs.
- `employee_worklog.py`: its `/submit` responses are counted.
- `generate_employee_performance_report._list_performance_teams`: team membership.
- `DATABASE_SCHEMA_MIGRATION_PROTOCOL.md`: records `usage_analytics.db`.

## Table Schema

`usage_analytics.db`. This is a standalone file beside the capacity DB (`/home/data/usage_analytics.db` on Azure), created automatically and gitignored.

| Column | Type | Constraint | Meaning |
|---|---|---|---|
| `id` | INTEGER | PRIMARY KEY AUTOINCREMENT | row id |
| `occurred_at_utc` | TEXT | NOT NULL | event time, ISO UTC |
| `event_date` | TEXT | NOT NULL | server-local calendar date (YYYY-MM-DD) |
| `account_id` | TEXT | NOT NULL DEFAULT '' | Atlassian account id |
| `display_name` | TEXT | NOT NULL DEFAULT '' | name at event time |
| `role` | TEXT | NOT NULL DEFAULT '' | admin / member |
| `event_type` | TEXT | NOT NULL | `page_view` or `worklog_submitted` |
| `path` | TEXT | NOT NULL DEFAULT '' | normalized page path (Referer page for worklogs) |

Indexes: `idx_usage_events_date(event_date)` and `idx_usage_events_user_date(account_id, display_name, event_date)`.

Also read: `assignee_hours_capacity.db.performance_teams` (`team_name`, `assignees_json`) and `epr_access/users.json`.

## Data Flow

1. A user opens any EPR page. `epr_auth` sets `g.epr_user`, the page is served, and `track_usage` inserts a `page_view` row.
2. A worklog submit succeeds, and `track_usage` inserts `worklog_submitted` with the Referer path.
3. An Admin opens `/settings/usage-analytics`, which loads `/api/page-categories` (titles) and then `GET /api/admin/usage-analytics?days=N`.
4. `build_usage_summary` aggregates `usage_events`, then joins teams and users.json in Python, and returns JSON. The page renders the tiles, chart and lists.

## Privacy

Only user id, name, role, date/time, event type and page path are stored. There are no query strings, report data, request bodies or IP addresses. The file is private server state like `worklog_requests/`: back it up securely and never commit or serve it.

## Change notes

- 2026-10-28: first release.
