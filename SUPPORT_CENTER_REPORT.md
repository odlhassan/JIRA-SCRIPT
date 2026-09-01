# Support Center Report

Shows **where** the R&D Technical Support Team's available hours are actually spent.
The Monthly Epic Plan Progress report already surfaces the 7-resource support team's
capacity/availability; this report answers the follow-up question: how many of those
hours were invested, on what, and how many support stories were resolved.

## Business Logic

- **Support set.** A standalone `support_center.db` stores every Jira issue whose
  **Work Type EPR** custom field (`customfield_10683`) equals **`Support`**. Issue types
  tagged can be Epic, Story, Subtask, or Bug Subtask. The canonical databases are **not**
  modified; this DB is joined back to canonical data on `issue_key` at query time.
- **Two kinds of Support stories** (Story-level issues in the support set):
  1. **Booking stories** — summary matches either:
     - Strict: `^\s*(?:technical\s+)?support\s+by\s+.+\(.+\)\s*$` (with parentheses),
       e.g. `Support by Nadeem (June 2026)`
     - Loose: `^\s*(?:technical\s+)?support\s+by\s+<name>\s+<Month>\s+<Year>$` (without
       parentheses), e.g. `Support by Abbas May 2026`, `Technical Support by Nadeem May 2026`
     These are roster entries; they show *who* is booked for support in a month. Their
     subtask worklogs still contribute to total invested hours.
  2. **Actual support stories** — any other support-tagged Story. These carry the real
     work: subtasks / bug subtasks with worklog hours.
- **Hours invested.** Total support work hours = sum of (a) actual-story subtask
  worklogs dated within the range, plus (b) booking-story subtask worklogs dated within
  the range. Per-subtask hours are counted from `canonical_worklogs.started_date` within
  the range. If no dated worklogs exist, falls back to
  `canonical_issue_actuals.total_worklog_hours`.
- **Date scoping.** Actual support stories are included if they **overlap** the date
  range via any of: (a) actual completion date in range, (b) start_date..due_date span
  overlaps range, (c) any subtask has a worklog with `started_date` within the range.
  This ensures in-progress stories with recent work are visible.
- **Resolved count.** Actual support stories within range whose status is Done/closed/
  resolved/complete/completed.
- **Hours available.** Reuses the Monthly Epic Plan capacity model
  (`build_workforce_month_payload` → `support_team.total_availability_hours`), summed
  across the calendar months overlapping the date range. It is **never** deduced from
  booking stories. If the capacity model can't be built, available hours read 0.
- **Utilization** = invested_hours / available_hours × 100.

## Business Cases

- R&D leadership wants visibility into how the support team's paid-for capacity is
  consumed across projects, separate from planned project delivery.
- Distinguishes "who is booked for support" (roster) from "what support work actually
  happened" (resolved stories + invested hours), so a booking does not get double-counted
  as effort.

## Examples

Support epic `O2-1` (June 2026) with:
- `O2-2` "Support by Nadeem (June 2026)" → booking story → roster row
  `{assignee: Nadeem, booked_for: June 2026}`.
- `O2-3` "Investigate flaky login" (Done, completes 2026-06-15) → actual support story.
  Subtasks `O2-3-1` (5h) + `O2-3-2` (3h) → invested 8h.
- `O2-4` "Performance tuning" (open, completes 2026-07-05) → actual support story, July.

Filter June 2026: support_story_count=1, resolved=1, invested=8h, roster=[Nadeem].
Filter July 2026: support_story_count=1, resolved=0, invested=9h.

## Explanations

When you open the report it defaults to the **current calendar month** (no global saved
filter is used). The **Birds-eye** view shows KPIs (available hours, invested hours,
utilization, resolved support stories, support member count), a by-project breakdown, and
the booking roster. Clicking a project name — or switching to the **Project drilldown**
tab and choosing a project — shows that project's actual support stories, each expandable
to its subtasks with logged hours and completion dates, plus that project's roster.

The project filter shows **full project names** (e.g. "Digital Log.", "Fintech Fuel")
sorted alphabetically. Jira project keys are shown in tooltips. Leadership can filter by
project without needing to know cryptic Jira keys. The Reset button also resets to the
current month rather than a blank date range.

## Front-end UI Fields

| Field | Type | Default | Controls |
|---|---|---|---|
| From / To | date inputs | **First/last day of current calendar month** | Date range matched against story actual completion date |
| Projects | multi-select | all | Full project names (alphabetically sorted); Jira key shown as tooltip. Restrict overview to selected projects; first selection drives drilldown |
| Apply | button | — | Re-fetch current view |
| Reset | button | — | Reload projects, reset dates to **current month**, return to Birds-eye |
| Birds-eye / Project drilldown | tabs | Birds-eye | Switch view |
| KPI: Hours Available | number | — | Capacity-derived availability across range months |
| KPI: Hours Invested | number | — | Sum of subtask worklogs of in-range stories |
| KPI: Utilization | percent | — | invested / available |
| KPI: Resolved Support Stories | number | — | Done actual support stories in range |
| KPI: Support Members | number | — | Capacity contributors |
| By project table | table | — | Per-project story/resolved/subtask counts + invested h/d; project name with key tooltip |
| Roster table | table | — | Booking stories: project name (key tooltip), assignee, booked-for month, story. Filtered to selected date range. |
| Project stories | expandable list | — | Actual support stories → subtasks (key, type, summary, assignee, status, logged h, completed) |

## Script Files

| File | Role |
|---|---|
| `support_center_sync.py` | Standalone Jira sync: fetch `customfield_10683 = Support`, write `support_center.db` |
| `support_center_service.py` | Data layer: classification, invested hours, resolved count, available hours, roster, project detail |
| `report_server.py` | Routes `/api/support-center/overview`, `/api/support-center/project/<key>`; report registration |
| `generate_support_center_report.py` | Publishes the HTML shell to root + `report_html/` |
| `support_center_report.html` | Static shell: filters + Birds-eye + Project drilldown, fetches the live API |
| `run_html_only.py` | Registers `support-center-html` generator step |
| `shared-nav.js` | Adds the Support Center nav item |
| `report_server.py` (page catalog) | `support_center_report` lives in `STATIC_REPORT_NAV_ITEMS`, so it is an assignable page on `/settings/page-categories` (user can pick a category for it) |
| `report_server.py` (SQL console) | `support_center` is a selectable database on `/settings/sql-console` (`_sql_console_target_path` + dropdown + sample queries) |
| `tests/test_support_center_service.py` | Focused unit tests |

### Page Categories

The Support Center report is registered as a report page (`page_key =
support_center_report`). On **`/settings/page-categories`** it appears in the page catalog
and a user can assign it to any page category; the assignment is stored in
`page_category_assignments` and drives the grouped left navigation.

### SQL Console

`support_center.db` is explorable on **`/settings/sql-console`** — pick **"Support Center
DB"** from the Database dropdown. Connections are opened read-only. Sample queries: list
tables, all support-tagged issues, counts by project, counts by type.

## Dependent & Impacted Files

- `canonical_report_data.py` — read-only loaders (`load_canonical_issues`,
  `load_canonical_worklogs`, `load_canonical_actuals_by_issue`, `resolve_canonical_run_id`).
  Not modified.
- `monthly_epic_plan_progress_service.py` — capacity reuse (`build_workforce_month_payload`,
  `HOURS_PER_DAY`, `_month_bounds`). Not modified; shares the support-team capacity model.
- `generate_employee_performance_report.py` — reads `support_team_config.members_json`
  read-only to show `Support team` chips for nested members in the Employee Performance
  Teams filter. It does not write the support roster or alter this schema.
- `db_schema_changelog.py` / `db_schema_changelog.db` — records the new `support_center.db`
  / `support_issues` table.
- `.gitignore` — `support_center.db` is a local-only artifact (rebuilt by the sync).
- `report_server.py` (`_run_canonical_phase1_refresh`, `generating_reports` stage) —
  **`support_center_sync.py` is now called here** so that every colossal refresh automatically
  repopulates `support_issues`. Previously missing, which caused the Support Center report to
  show empty data after a colossal refresh. In `_run_canonical_compute()` a non-zero exit from
  this script fails the whole Compute run, so its DB path must be writable.
- `report_output_paths.py` — `is_writable_directory()` backs
  `support_center_sync._writable_db_path()`. When the directory holding `support_center.db` is
  read-only (Azure `WEBSITE_RUN_FROM_PACKAGE`), the sync uses `$HOME/data/support_center.db`
  instead, seeding it once from the deployed copy so packaged support data is preserved.
  Set `SUPPORT_CENTER_DB_PATH=/home/data/support_center.db` to control this explicitly.
  See `AZURE_APP_SERVICE.md`.

## Table Schema

### `support_center.db` → `support_issues` (written by this module)

| Column | Type | Constraint | Meaning |
|---|---|---|---|
| issue_key | TEXT | PRIMARY KEY | Jira issue key tagged Work Type EPR = Support |
| project_key | TEXT | NOT NULL '' | Project key |
| issue_type | TEXT | NOT NULL '' | Epic / Story / Sub-task / Bug Subtask |
| summary | TEXT | NOT NULL '' | Issue summary (used for booking-story regex) |
| work_type_value | TEXT | NOT NULL '' | Extracted Work Type EPR option (Support) |
| synced_at_utc | TEXT | NOT NULL '' | Sync timestamp |

### Canonical tables read read-only (owned elsewhere)

- `canonical_issues` — issue_key, project_key, issue_type, summary, status, assignee,
  start_date, due_date, total_hours_logged, story_key, epic_key.
- `canonical_worklogs` — issue_key, hours_logged, started_date.
- `canonical_issue_actuals` — issue_key, last_worklog_date, actual_complete_date,
  total_worklog_hours.
- `canonical_refresh_state` — last_success_run_id (run scoping).

## Data Flow

1. **Sync** (`support_center_sync.sync`): JQL `cf[10683] = "Support"` paged via
   `/rest/api/3/search/jql` → full-replace `support_issues` in `support_center.db`.
   This runs automatically at the end of every **colossal refresh** (`_run_canonical_phase1_refresh`
   → `generating_reports` stage), or on demand via the Support Center report's **Refresh** button
   (which invokes `REPORT_REFRESH_CHAINS["support_center"]`), or manually with
   `python support_center_sync.py`.
2. **Request**: browser opens `support_center_report.html`, loads `/api/projects`, then
   GETs `/api/support-center/overview?from=&to=&projects=`.
3. **Service** (`build_support_center_overview`): resolve canonical run id → load issues /
   actuals / worklogs (read-only) + support keys → classify booking vs actual stories →
   filter actual stories by date overlap (completion in range OR start/due span overlaps
   OR subtask worklogs in range) → sum subtask hours within the date range → also sum
   booking-story subtask hours within range → compute resolved count → sum available hours
   from capacity model → return KPIs + by_project + roster.
4. **Drilldown**: `/api/support-center/project/<key>` →
   `build_support_center_project_detail` returns per-story subtask detail + roster.
5. **Render**: HTML shell paints KPIs, tables, and expandable stories.

## Verification

- Sync: `python support_center_sync.py` populates `support_center.db` (needs live Jira).
- API: `GET /api/support-center/overview?from=2026-06-01&to=2026-06-30`.
- UI: `http://127.0.0.1:3000/support_center_report.html`.
- Tests: `python -m pytest tests/test_support_center_service.py -q`.
