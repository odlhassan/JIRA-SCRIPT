1. # Support Hour Bookings (Performance Settings)

## Screen

- Name: Support Hour Bookings
- Route: `/settings/performance` (new section on the existing Performance Point Settings page)
- Purpose: Give the admin a monthly control to create Technical Support Team hour "bookings" per member and split each member's booked hours across projects by percentage. The destination report(s) that will consume these bookings is a separate, later decision — this screen only captures the configuration.

## Business Logic

For a selected `booking_month` (`YYYY-MM`) and a chosen `capacity_profile_key` (an existing saved profile from `/settings/capacity`, keyed as `from_date|to_date`):

1. **System capacity hours** per team member = `workdays_in_month(profile) x hours_per_day(profile)`.
   - Workdays are Monday–Friday.
   - Days inside the profile's `holiday_dates_json` are excluded entirely.
   - Days inside the profile's Ramadan range (if set) use `ramadan_hours_per_day` instead of `standard_hours_per_day`.
   - Example: July 2026 has 23 weekdays, no holidays; at 8h/day → `23 x 8 = 184.0` hours.
2. **Leave hours** default to an admin-configurable assumption (default `16.0`, i.e. ~2 leave days x 8h) because actual planned leave is rarely entered for a future month yet. The admin can override this per member at any time (e.g. once real leave gets planned in the RLT tracker).
3. **Availability hours** = `system_capacity_hours - leave_hours` (floored at 0). This is auto-computed and shown read-only, but is only a *reference* number.
4. **Booking hours** is the number the admin actually types for that member for the month — it defaults to `availability_hours` when the row is first created, but the admin is free to type any other number (e.g. a resource who is only half on support that month: `Booking Hours = 84` instead of the full `168` availability).
5. **Project allocation percentages**: for each active project, the admin types a fraction between `0.0` and `1.0` (e.g. `0.3` = 30%). The system computes `hours_for_project = booking_hours x percentage` on every read — this value is never stored, so it can never drift from the underlying booking/percentage inputs.
6. **Validation guards**:
   - Percentage inputs must be a fraction (`0.0`–`2.0` accepted so rounding never blocks a save, but a value like `30` typed in place of `0.3` is rejected with a clear error).
   - A row is flagged `over_allocated` if the sum of a member's percentages exceeds `100%` (`1.0`), and `over_capacity` if `booking_hours` exceeds `system_capacity_hours`. Both are warnings, not hard blocks, since partial-month / partial-role bookings are valid.

## Business Cases

- The admin needs to plan, each month, how many hours each Technical Support Team member is expected to spend on support work, and how those hours split across the projects the team supports (OmniConnect, OmniChat, Fintech Fuel, Digital Log, etc.).
- At the start of a month, actual leave is rarely planned yet, so the tool assumes a standard "2 leave days" buffer per person and lets the admin override it once real numbers are known, or when someone is only partially available for support that month.
- Whoever consumes this configuration downstream (a report or another tool) can read `booking_hours` and each project's computed hours directly, instead of re-deriving them from raw percentages every time.

## Examples

Input: July 2026 (23 weekdays, no holidays, profile `2026-01-01|2026-12-31` @ 8h/day):

| Member | Booking (h) | OmniConnect % | OmniChat % | Fintech Fuel % | Digital Log % |
| --- | --- | --- | --- | --- | --- |
| Abbas | 84 | 0.3 | 0.3 | 0.2 | 0.2 |
| Ameer | 168 | 0.5 | — | 0.5 | — |

Computed mirror hours (read-only, `booking_hours x percentage`):

| Member | OmniConnect (h) | OmniChat (h) | Fintech Fuel (h) | Digital Log (h) |
| --- | --- | --- | --- | --- |
| Abbas | 25.2 | 25.2 | 16.8 | 16.8 |
| Ameer | 84.0 | 0.0 | 84.0 | 0.0 |

System capacity for both members that month = `23 x 8 = 184.0`; with the default `16.0` leave assumption, availability = `168.0`. Ameer keeps the full `168` as his booking; Abbas is manually reduced to `84` because he's only half on support that month.

## Explanations

Har month ke shuru mein admin ek month aur capacity profile chunta hai. System khud calculate kar deta hai keh us profile ke hisaab se har support team member ki total capacity kitni hai (working days x hours/day). Chunke month ke start mein leaves plan nahi hoti, system ek default assumption (2 din/16 ghantay) laga kar "availability" dikhata hai — admin chahe tou ye number khud badal sakta hai, ya kisi banday ki booking hours manually kam/zyada likh sakta hai (jaisay koi half support aur half kisi aur kaam per ho). Uske baad admin har project ke against percentage likhta hai, aur system foran hours ki mirror matrix bana kar dikhata hai jise admin kahin bhi copy kar sakta hai.

## Front-end UI Fields

Section id: `#support-bookings` on `/settings/performance`.

| Field | Type | Required | Default | Description | Validation / Rules |
| --- | --- | --- | --- | --- | --- |
| Month (`shb-month`) | Month picker | Yes | Current month | Selects the `booking_month` (`YYYY-MM`) to view/edit. | Must be a valid month. |
| Capacity profile (`shb-profile`) | Dropdown | Yes (for Init) | First saved profile | Lists saved capacity profiles as `from_date → to_date (hours/day)`. | Populated from `GET /api/performance/support-bookings/capacity-profiles`. |
| Assumed leave hours (`shb-default-leave`) | Number | No | `16` | Default leave-hours assumption applied only to *newly created* member rows for the month. | Must be `>= 0`. |
| Load / Initialize month (`shb-load-btn`) | Button | No | Enabled | Creates header rows for any support team member missing from the selected month (existing rows/edits untouched), then loads the full matrix. | Calls `POST /api/performance/support-bookings/init-month`. |
| Team member row — Capacity (h) | Read-only | N/A | Computed | System-computed capacity for the month. | Not editable directly; changes with capacity profile. |
| Team member row — Leave (h) | Number input | No | `16` (or existing value) | Assumed/actual planned leave hours for that member/month. | `>= 0`; recomputes `availability_hours` on save. |
| Team member row — Availability (h) | Read-only | N/A | Computed | `capacity - leave`, floored at 0. | Not directly editable. |
| Team member row — Booking (h) | Number input | Yes | `availability_hours` on creation | The hours the admin books that member for support work this month. | `>= 0`; flags `over_capacity` styling if it exceeds system capacity. |
| Per-project `% ` cell | Number input | No | `0` | Fraction of the member's booking hours allocated to that project. | `0.0`–`2.0`; row highlighted if the member's total exceeds `1.0`. |
| Per-project `(h)` cell | Read-only, click-to-copy | N/A | Computed | `booking_hours x percentage`, click copies the value to clipboard. | Recomputed on every load/save; never stored. |
| Status line (`shb-status`) | Read-only text | No | Empty | Feedback for load/save/init actions (success/error). | — |

## Script Files

- `support_booking_registry.py` — new module: schema init, capacity/leave calculation, header + allocation CRUD, and the `get_month_matrix()` read model that returns the computed hours mirror matrix.
- `report_server.py` — new API routes under `/api/performance/support-bookings*`, plus the new `#support-bookings` section (HTML/CSS/JS) appended to `_performance_settings_html()`.
- `tests/test_support_booking_registry.py` — unit tests for month-bounds capacity math, header init/override, allocation upserts, the mirror-matrix computation, and over-allocation/over-capacity flags.
- `manage_fields_registry.py`, `managed_projects_registry.py` — existing sibling modules whose CRUD conventions (`init_*_db`, normalize/validate, `created_at_utc`/`updated_at_utc`) this module follows.

## Dependent & Impacted Files

- `generate_employee_performance_report.py` — owns `_load_support_team_members()`, the source of truth for which names appear as rows in this screen (`support_team_config` table). Any change to how the support team roster is edited also affects this screen's row list.
- `generate_assignee_hours_report.py` — owns the capacity-profile CRUD (`assignee_capacity_settings`) that this screen's capacity dropdown and system-capacity math read from.
- `managed_projects_registry.py` / `/api/projects` — supplies the active project list used as crosstab columns; if a project is deactivated, previously-saved allocations for it still surface (so no silent data loss), but it will no longer appear for new allocation entry unless reactivated.
- **No report currently reads `support_hour_booking_headers` / `support_hour_booking_allocations`.** This is intentional — which report(s) will consume these bookings is a decision deferred to a future task. Whichever report is chosen next should be added to this section once wired up.

## Table Schema

Database: `assignee_hours_capacity.db`

### `support_hour_booking_headers`

| Column | Type | Description |
| --- | --- | --- |
| `id` | INTEGER PK | Autoincrement row id. |
| `booking_month` | TEXT | `YYYY-MM`. |
| `team_member` | TEXT | Must match a name from `support_team_config.members_json`. |
| `capacity_profile_key` | TEXT | `from_date\|to_date` of the `assignee_capacity_settings` row used for the capacity calculation. |
| `system_capacity_hours` | REAL | Auto-computed workdays x hours/day for the month, per the chosen profile. |
| `leave_hours` | REAL | Assumed/actual planned leave hours for the member that month (admin-editable). |
| `availability_hours` | REAL | `system_capacity_hours - leave_hours`, floored at 0. |
| `booking_hours` | REAL | The admin-entered support booking for the member/month (defaults to `availability_hours`). |
| `notes` | TEXT | Free-text notes (reserved for future use; not yet exposed in the UI). |
| `created_at_utc` / `updated_at_utc` | TEXT | ISO-ish UTC timestamps. |

Unique constraint: `(booking_month, team_member)`.

### `support_hour_booking_allocations`

| Column | Type | Description |
| --- | --- | --- |
| `id` | INTEGER PK | Autoincrement row id. |
| `booking_month` | TEXT | `YYYY-MM`. |
| `team_member` | TEXT | Same value as in `support_hour_booking_headers.team_member`. |
| `project_key` | TEXT | Uppercased project key (matches `managed_projects.project_key`). |
| `percentage` | REAL | Fraction (0.0–1.0, occasionally slightly above for rounding) of the member's `booking_hours` allocated to this project. |
| `updated_at_utc` | TEXT | ISO-ish UTC timestamp. |

Unique constraint: `(booking_month, team_member, project_key)`. A row is deleted automatically when its `percentage` is saved as `0`.

## Data Flow

1. Admin opens `/settings/performance` and scrolls to the **Support Hour Bookings** section.
2. On page load, the UI calls `GET /api/performance/support-bookings/capacity-profiles` to populate the profile dropdown, and `GET /api/performance/support-bookings?month=<current month>` to show any existing data.
3. Admin picks a month + capacity profile and clicks **Load / Initialize month** → `POST /api/performance/support-bookings/init-month`. This reads `support_team_config` for the current roster, computes `system_capacity_hours` per member via `compute_person_month_capacity_hours()` (reading `assignee_capacity_settings`), and inserts a header row per member missing from that month (existing rows are left untouched).
4. Editing **Leave (h)** or **Booking (h)** on a row triggers `PUT /api/performance/support-bookings/<team_member>`, which recomputes `availability_hours` server-side and persists the row.
5. Editing a project `%` cell triggers `PUT /api/performance/support-bookings/<team_member>/allocations`, which upserts/deletes the relevant `support_hour_booking_allocations` rows and returns the refreshed matrix (including computed hours) so the UI can re-render instantly.
6. `GET /api/performance/support-bookings?month=...` (used for both initial load and refreshes) calls `get_month_matrix()`, which joins headers + allocations + the active project list from `managed_projects` and returns, per member, `allocations` (%), `hours` (computed), `allocation_pct_total`, `over_allocated`, and `over_capacity`.
7. The right-hand "(h)" cells are click-to-copy so the admin can paste any computed value elsewhere. No destination report reads this data yet — that wiring is a separate future task.
