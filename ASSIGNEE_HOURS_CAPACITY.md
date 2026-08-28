# Assignee Hours Capacity Backend

This report now supports a SQLite-backed Capacity Planning form in the HTML report.

## User Guide Docs

Detailed functional documentation is available at:

- Module overview: `docs/capacity-user-guide/00-capacity-overview.md`
- Capacity settings page fields and behavior: `docs/capacity-user-guide/screens/01-capacity-settings-page.md`
- Assignee Hours integration and KPI logic: `docs/capacity-user-guide/screens/02-assignee-hours-capacity-integration.md`
- Nested View linkage and profile application flow: `docs/capacity-user-guide/screens/03-nested-view-capacity-integration.md`
- Unified cross-report info-drawer and logic docs: `docs/report-user-guide/00-report-overview.md`

## Files Updated

- Python generator and server:
  - `generate_assignee_hours_report.py`
- HTML outputs:
  - `assignee_hours_report.html`
  - `report_html/assignee_hours_report.html`
- Excel output:
  - `assignee_hours_report.xlsx`
- SQLite storage:
  - `assignee_hours_capacity.db`
  - `assignee_hours_capacity.sqlite`

### Dependent & Impacted Files

- `support_center_service.py` reads `assignee_hours_capacity.db` **read-only** (canonical
  issues/worklogs/actuals + support-team capacity) to power the Support Center report. It
  never writes to or alters this DB; its own data lives in the standalone
  `support_center.db`. See `SUPPORT_CENTER_REPORT.md`.
- `generate_employee_performance_report.py` reads `performance_teams`,
  `performance_resource_resignations`, and `support_team_config` **read-only** to render the
  custom Employee Performance Teams filter with nested members, resignation status, and
  `Support team` chips.
- `rnd_muscle_utilization_service.py` now stores feature-owned configuration and planner
  state in the standalone `rnd_muscle_utilization.db`. It reads this capacity DB only as
  the source for Epics Planner rows and canonical assignee/worklog names.

## Run Modes

- Static generation:
  - `python generate_assignee_hours_report.py`
  - The generated HTML now embeds leave daily rows from `rlt_leave_report.xlsx` and can compute leave KPIs without backend API.
- Server mode (required for persistent capacity save/load):
  - `python generate_assignee_hours_report.py --server --port 5000`
  - Open `http://localhost:5000`

## Capacity APIs

- `GET /api/capacity?from=YYYY-MM-DD&to=YYYY-MM-DD`
- `POST /api/capacity`
- `POST /api/capacity/calculate`
- `GET /api/capacity/profiles`

## Business Logic

- The capacity database path is resolved by `report_server.py` from `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, defaulting to `assignee_hours_capacity.db` under the app root for local development.
- RnD Muscle Utilization has its own database path, resolved from `JIRA_RND_MUSCLE_UTILIZATION_DB_PATH`, defaulting to `rnd_muscle_utilization.db` under the app root.
- Runtime path values are trimmed and accidental wrapping quotes are removed before `Path` resolution, so Azure app settings like `"/home/data/assignee_hours_capacity.db"` still resolve to `/home/data/assignee_hours_capacity.db`.
- The capacity DB parent directory is created before the first SQLite connection in `_init_capacity_db`, which prevents cold-start failure when `/home/data` or a nested local test directory exists only as a configured path.
- On Azure, an unwritable configured path falls back to `$HOME/data/assignee_hours_capacity.db` with a stderr warning so the web worker can boot instead of returning the generic App Service Application Error page.

## Business Cases

- Production operators keep mutable planning and settings data outside the deployed app folder so code deploys do not overwrite the live SQLite database.
- Delivery leads need Capacity Settings, Team Capacity Planner, Employee Performance, and refresh screens to remain available immediately after an Azure restart, even if the deployment instance did not preserve an Oryx virtual environment.
- Developers need local tests to create isolated nested DB paths without manually creating every parent folder.

## Examples

| Input | Output |
| --- | --- |
| `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db` | Server creates `/home/data` if needed and opens `/home/data/assignee_hours_capacity.db`. |
| `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH="/home/data/assignee_hours_capacity.db"` | Quotes are stripped; SQLite opens `/home/data/assignee_hours_capacity.db`. |
| Local test DB path `tmp/nested/capacity/assignee_hours_capacity.db` | `_init_capacity_db` creates `tmp/nested/capacity` before creating the DB. |

## Explanations

At server startup, `wsgi.py` calls `create_report_server_app()`. The server resolves the capacity DB path, verifies that the parent folder can be written, initializes capacity tables, then initializes dependent settings tables used by Employee Performance, Team Capacity Planner, managed projects, page categories, canonical refresh, and report refresh APIs. RnD Muscle Utilization resolves a separate SQLite path for its own `rnd_muscle_*` tables and uses the capacity DB as a read source for Epics Planner/canonical data. If the configured production capacity path is unusable, the server logs the fallback path and continues with `$HOME/data/assignee_hours_capacity.db` rather than failing worker boot.

Canonical refresh lifecycle state is also stored here. `canonical_fetch_runs` records raw Jira Fetch snapshots and checkpoints, while `canonical_compute_runs` records the downstream precomputation that makes a Fetch visible to shared reports. `employee_performance_scoped_runs` records the separate assignee-specific Employee Performance flow; it must not change the global promoted canonical run.

## Front-end UI Fields

| Field | Type | Default | Valid values / range | Controls | Example |
| --- | --- | --- | --- | --- | --- |
| Capacity Settings date range | date inputs | report-selected range | ISO dates | Key for saved capacity profile rows | `2026-01-01` to `2026-12-31`. |
| Employee count | numeric input | configured profile value | non-negative integer | Multiplier for gross capacity | `30`. |
| Standard hours/day | numeric input | configured profile value | non-negative hours | Normal working-day capacity | `8`. |
| Ramadan start/end | date inputs | blank | ISO dates or blank | Date range where Ramadan hours apply | `2026-02-18` to `2026-03-19`. |
| Ramadan hours/day | numeric input | configured profile value | non-negative hours | Working-day capacity during Ramadan | `6`. |
| Holiday dates | date list/input | blank | ISO dates | Workdays excluded from capacity | `2026-03-23`. |
| Save Capacity | button | enabled when form valid | click command | Writes or updates the selected profile | Save annual 2026 profile. |
| Reuse Saved Capacity | selector/button | no profile selected | saved ranges | Loads one saved profile into the current range | Reuse `2026-01-01 - 2026-12-31`. |

## Script Files

| File | Role |
| --- | --- |
| `generate_assignee_hours_report.py` | Capacity formulas, SQLite schema initialization, capacity APIs for standalone server mode, and Assignee Hours generation. |
| `report_server.py` | Production/local Flask app, shared capacity DB path resolution, startup schema initialization, and settings/report APIs. |
| `startup.txt` | Azure shell startup script that imports vendored Python dependencies and starts `wsgi:app`. |
| `.github/workflows/azure-appservice-deploy.yml` | Builds the deploy ZIP, vendors `requirements.txt` into `.python_packages`, and marks `startup.txt` executable before packaging. |
| `ASSIGNEE_HOURS_CAPACITY.md` | Operational and business documentation for capacity storage and behavior. |

## Table Schema

### `assignee_capacity_settings`

| Column | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| `id` | INTEGER | PRIMARY KEY AUTOINCREMENT | Internal row id. |
| `from_date` | TEXT | NOT NULL, unique pair with `to_date` | Profile start date. |
| `to_date` | TEXT | NOT NULL, unique pair with `from_date` | Profile end date. |
| `employee_count` | INTEGER | NOT NULL | Number of employees represented by the profile. |
| `standard_hours_per_day` | REAL | NOT NULL | Normal working hours per employee per day. |
| `ramadan_start_date` | TEXT | nullable | Optional Ramadan period start. |
| `ramadan_end_date` | TEXT | nullable | Optional Ramadan period end. |
| `ramadan_hours_per_day` | REAL | NOT NULL | Daily hours inside Ramadan range. |
| `holiday_dates_json` | TEXT | NOT NULL | JSON array of holiday ISO dates to exclude. |
| `created_at_utc` | TEXT | NOT NULL | Creation timestamp. |
| `updated_at_utc` | TEXT | NOT NULL | Last update timestamp. |

### `team_capacity_planned_assignments`

| Column | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| `id` | INTEGER | PRIMARY KEY AUTOINCREMENT | Internal row id. |
| `issue_key` | TEXT | NOT NULL, indexed | Jira issue key for planned assignment tracking. |
| `assignee_display_name` | TEXT | NOT NULL DEFAULT '' | Human-readable assignee name. |
| `assignee_account_id` | TEXT | NOT NULL DEFAULT '' | Jira account id when available. |
| `jira_synced` | INTEGER | NOT NULL DEFAULT 0 | Whether the assignment synced back to Jira. |
| `jira_error` | TEXT | NOT NULL DEFAULT '' | Last sync error message. |
| `created_at_utc` | TEXT | NOT NULL DEFAULT '' | Creation timestamp. |
| `updated_at_utc` | TEXT | NOT NULL DEFAULT '' | Last update timestamp. |

## Data Flow

1. Azure or local startup calls `wsgi.py` / `run_server.py`.
2. `report_server.py` normalizes `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, resolves relative paths against the app root, creates the parent folder, and validates write access.
3. `_init_capacity_db()` creates capacity tables before dependent modules initialize their own tables in the same SQLite database.
4. Capacity Settings writes saved profiles through the capacity APIs.
5. Assignee Hours, Employee Performance, Team Capacity Planner, Nested View, canonical refresh, and settings routes read the shared DB path for profile, team, leave, managed-project, and refresh state.
6. Reports return HTML/API responses to the browser while SQLite state remains in the resolved persistent DB file.

## Reuse Saved Capacity

- Capacity settings are saved per date range.
- In the Capacity Planning form, use **Reuse Saved Capacity** to select a previously saved range.
- Click **Apply Profile To Current Range** to load those settings into the current selected report range.
- Then click **Save Capacity** to persist that profile for the current range.

## Capacity Formula

`available_capacity_hours = employees * (non_ramadan_workdays * standard_hours_per_day + ramadan_workdays * ramadan_hours_per_day)`

Where:

- Working days are Monday-Friday in selected date range
- Holiday dates are excluded
- Ramadan is a contiguous start/end range if provided

## Capacity Planning KPIs

- `Total Capacity`:
  - Raw profile capacity (`available_capacity_hours`)
- `Leave Hours`:
  - `planned_taken_hours + planned_not_taken_hours + unplanned_taken_hours`
- `Remaining Capacity`:
  - `Total Capacity - Leave Hours`
- `Project Planned Hours`:
  - Existing planned-minus-leaves value (caption updated)
- `Project Actual Hours`:
  - Total logged project hours excluding project key `RLT` (case-insensitive)
- `Project Plan - Actual Hours`:
  - `Project Planned Hours - Project Actual Hours`

## Capacity Subtraction Card

`Capacity After Leaves = Available Capacity - Project Actual Hours - Leave Hours`

## API Leave Metrics

`/api/capacity`, `/api/capacity/calculate`, and `POST /api/capacity` include:

- `planned_taken_hours`
- `planned_not_taken_hours`
- `unplanned_taken_hours`
- `taken_hours`
- `not_yet_taken_hours`
- `taken_days`
- `not_yet_taken_days`
- `remaining_balance_hours`
- `remaining_balance_days`
