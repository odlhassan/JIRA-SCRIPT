"""Support Hour Bookings registry.

Gives Performance Settings admins a monthly control to:
  1. Pick a month + a saved capacity profile (from assignee_capacity_settings).
  2. See each support team member's system capacity hours for that month
     (workdays x hours/day, honoring Ramadan/holiday rules from the profile).
  3. Override the "available" hours per member (leave hours are rarely
     planned at the start of a month, so the admin can type an assumed
     leave/availability figure, or a manual override for someone who is
     only half on support that month).
  4. Allocate that member's booking hours across projects using percentages;
     the module returns the computed hours-per-project mirror matrix so the
     admin can copy any cell value straight into whichever report/tool needs
     it (the actual destination report is decided separately).
"""

from __future__ import annotations

import calendar
import json
import re
import sqlite3
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

DEFAULT_ASSUMED_LEAVE_HOURS = 16.0  # ~2 leave days x 8h/day, per admin's stated default assumption
PREFERRED_SUPPORT_BOOKING_PROJECT_NAMES = (
    "OmniConnect",
    "OmniChat",
    "Fintech Fuel",
    "Digital Log",
    "ODL Miscellaneous",
)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: object, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def normalize_booking_month(value: object) -> str:
    text = _to_text(value)
    if not re.match(r"^\d{4}-\d{2}$", text):
        raise ValueError("booking_month must be in 'YYYY-MM' format.")
    year, month = int(text[:4]), int(text[5:7])
    if month < 1 or month > 12:
        raise ValueError("booking_month must have a valid month (01-12).")
    return f"{year:04d}-{month:02d}"


def _month_bounds(booking_month: str) -> tuple[date, date]:
    year, month = int(booking_month[:4]), int(booking_month[5:7])
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    return first, last


def normalize_team_member(value: object) -> str:
    name = _to_text(value)
    if not name:
        raise ValueError("team_member is required.")
    return name


def normalize_project_key(value: object) -> str:
    key = _to_text(value).upper()
    if not key:
        raise ValueError("project_key is required.")
    return key


def support_booking_project_label(project: dict | object) -> str:
    if not isinstance(project, dict):
        return ""
    for field in ("display_name", "project_name", "project_key"):
        label = _to_text(project.get(field))
        if label:
            return label
    return ""


def sort_support_booking_projects(projects: list[dict]) -> list[dict]:
    priority_map = {name.casefold(): idx for idx, name in enumerate(PREFERRED_SUPPORT_BOOKING_PROJECT_NAMES)}

    def sort_key(project: dict) -> tuple[int, int, str, str]:
        label = support_booking_project_label(project)
        project_key = _to_text(project.get("project_key")).upper()
        rank = priority_map.get(label.casefold())
        if rank is None:
            rank = priority_map.get(project_key.casefold())
        if rank is None:
            return (1, len(priority_map), label.casefold(), project_key)
        return (0, rank, label.casefold(), project_key)

    return sorted(projects, key=sort_key)


def normalize_percentage(value: object) -> float:
    pct = _to_float(value, 0.0)
    if pct < 0:
        raise ValueError("percentage must be >= 0.")
    if pct > 2:
        # Guard against accidental "30" instead of "0.3" entries; percentages
        # are stored as fractions (0.0 - 1.0, occasionally slightly over for
        # rounding), never as raw 0-100 integers.
        raise ValueError("percentage must be expressed as a fraction (0.0 - 1.0), not a whole percent.")
    return round(pct, 4)


def _load_support_team_members(db_path: Path) -> list[str]:
    if not db_path.exists():
        return []
    try:
        with sqlite3.connect(db_path) as conn:
            table_row = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'support_team_config'"
            ).fetchone()
            if not table_row:
                return []
            row = conn.execute(
                "SELECT members_json FROM support_team_config WHERE key = 'members'"
            ).fetchone()
    except sqlite3.Error:
        return []
    if not row:
        return []
    try:
        parsed = json.loads(_to_text(row[0]) or "[]")
    except json.JSONDecodeError:
        return []
    return sorted({_to_text(name) for name in parsed if _to_text(name)}, key=lambda s: s.casefold())


def _load_capacity_profile(db_path: Path, capacity_profile_key: str) -> dict | None:
    if "|" not in capacity_profile_key:
        return None
    from_date, to_date = capacity_profile_key.split("|", 1)
    if not db_path.exists():
        return None
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT from_date, to_date, standard_hours_per_day,
                   ramadan_start_date, ramadan_end_date, ramadan_hours_per_day,
                   holiday_dates_json
            FROM assignee_capacity_settings
            WHERE from_date = ? AND to_date = ?
            """,
            (from_date, to_date),
        ).fetchone()
    except sqlite3.Error:
        return None
    finally:
        conn.close()
    if not row:
        return None
    try:
        holidays = json.loads(_to_text(row["holiday_dates_json"]) or "[]")
    except json.JSONDecodeError:
        holidays = []
    return {
        "from_date": _to_text(row["from_date"]),
        "to_date": _to_text(row["to_date"]),
        "standard_hours_per_day": _to_float(row["standard_hours_per_day"], 8.0),
        "ramadan_start_date": _to_text(row["ramadan_start_date"]),
        "ramadan_end_date": _to_text(row["ramadan_end_date"]),
        "ramadan_hours_per_day": _to_float(row["ramadan_hours_per_day"], 6.5),
        "holiday_dates": holidays,
    }


def compute_person_month_capacity_hours(db_path: Path, booking_month: str, capacity_profile_key: str) -> float:
    """Hours a single person can work in the given month per the chosen capacity profile.

    Workdays (Mon-Fri) minus holidays; Ramadan days (if the profile's Ramadan
    range overlaps the month) use ramadan_hours_per_day instead of the
    standard hours/day.
    """
    profile = _load_capacity_profile(db_path, capacity_profile_key)
    if not profile:
        return 0.0

    month_first, month_last = _month_bounds(booking_month)
    profile_from = date.fromisoformat(profile["from_date"])
    profile_to = date.fromisoformat(profile["to_date"])
    range_start = max(month_first, profile_from)
    range_end = min(month_last, profile_to)
    if range_end < range_start:
        return 0.0

    holiday_set = set()
    for item in profile["holiday_dates"]:
        try:
            holiday_set.add(date.fromisoformat(_to_text(item)))
        except ValueError:
            continue

    ramadan_start = None
    ramadan_end = None
    if profile["ramadan_start_date"] and profile["ramadan_end_date"]:
        try:
            ramadan_start = date.fromisoformat(profile["ramadan_start_date"])
            ramadan_end = date.fromisoformat(profile["ramadan_end_date"])
        except ValueError:
            ramadan_start = ramadan_end = None

    total_hours = 0.0
    cursor = range_start
    while cursor <= range_end:
        if cursor.weekday() < 5 and cursor not in holiday_set:
            in_ramadan = bool(ramadan_start and ramadan_end and ramadan_start <= cursor <= ramadan_end)
            total_hours += profile["ramadan_hours_per_day"] if in_ramadan else profile["standard_hours_per_day"]
        cursor += timedelta(days=1)
    return round(total_hours, 2)


def list_capacity_profile_options(db_path: Path) -> list[dict]:
    if not db_path.exists():
        return []
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT from_date, to_date, employee_count, standard_hours_per_day, updated_at_utc
            FROM assignee_capacity_settings
            ORDER BY updated_at_utc DESC, from_date DESC
            """
        ).fetchall()
    except sqlite3.Error:
        return []
    finally:
        conn.close()
    out: list[dict] = []
    for row in rows:
        out.append(
            {
                "capacity_profile_key": f"{row['from_date']}|{row['to_date']}",
                "from_date": _to_text(row["from_date"]),
                "to_date": _to_text(row["to_date"]),
                "employee_count": int(row["employee_count"] or 0),
                "standard_hours_per_day": _to_float(row["standard_hours_per_day"], 8.0),
            }
        )
    return out


def init_support_booking_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS support_hour_booking_headers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                booking_month TEXT NOT NULL,
                team_member TEXT NOT NULL,
                capacity_profile_key TEXT NOT NULL DEFAULT '',
                system_capacity_hours REAL NOT NULL DEFAULT 0,
                leave_hours REAL NOT NULL DEFAULT 0,
                availability_hours REAL NOT NULL DEFAULT 0,
                booking_hours REAL NOT NULL DEFAULT 0,
                notes TEXT NOT NULL DEFAULT '',
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL,
                UNIQUE(booking_month, team_member)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS support_hour_booking_allocations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                booking_month TEXT NOT NULL,
                team_member TEXT NOT NULL,
                project_key TEXT NOT NULL,
                percentage REAL NOT NULL DEFAULT 0,
                updated_at_utc TEXT NOT NULL,
                UNIQUE(booking_month, team_member, project_key)
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_shb_headers_month ON support_hour_booking_headers(booking_month)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_shb_alloc_month ON support_hour_booking_allocations(booking_month)"
        )
        conn.commit()
    finally:
        conn.close()


def _row_to_header(row: sqlite3.Row) -> dict:
    return {
        "id": int(row["id"]),
        "booking_month": _to_text(row["booking_month"]),
        "team_member": _to_text(row["team_member"]),
        "capacity_profile_key": _to_text(row["capacity_profile_key"]),
        "system_capacity_hours": _to_float(row["system_capacity_hours"]),
        "leave_hours": _to_float(row["leave_hours"]),
        "availability_hours": _to_float(row["availability_hours"]),
        "booking_hours": _to_float(row["booking_hours"]),
        "notes": _to_text(row["notes"]),
        "created_at_utc": _to_text(row["created_at_utc"]),
        "updated_at_utc": _to_text(row["updated_at_utc"]),
    }


def init_month_bookings(
    db_path: Path,
    booking_month: str,
    capacity_profile_key: str,
    default_leave_hours: float = DEFAULT_ASSUMED_LEAVE_HOURS,
) -> list[dict]:
    """Create header rows (if missing) for every current support team member for a month.

    Existing rows for the month are left untouched. Returns all header rows
    for the month after initialization.
    """
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    members = _load_support_team_members(db_path)
    now = _utc_now_iso()
    system_capacity_hours = compute_person_month_capacity_hours(db_path, month, capacity_profile_key)

    conn = sqlite3.connect(db_path)
    try:
        existing_members = {
            _to_text(r[0])
            for r in conn.execute(
                "SELECT team_member FROM support_hour_booking_headers WHERE booking_month = ?",
                (month,),
            ).fetchall()
        }
        for member in members:
            if member in existing_members:
                continue
            availability_hours = round(max(system_capacity_hours - default_leave_hours, 0.0), 2)
            conn.execute(
                """
                INSERT INTO support_hour_booking_headers (
                    booking_month, team_member, capacity_profile_key,
                    system_capacity_hours, leave_hours, availability_hours, booking_hours,
                    notes, created_at_utc, updated_at_utc
                ) VALUES (?, ?, ?, ?, ?, ?, ?, '', ?, ?)
                """,
                (
                    month,
                    member,
                    capacity_profile_key,
                    system_capacity_hours,
                    default_leave_hours,
                    availability_hours,
                    availability_hours,
                    now,
                    now,
                ),
            )
        conn.commit()
    finally:
        conn.close()
    return list_month_headers(db_path, month)


def list_month_headers(db_path: Path, booking_month: str) -> list[dict]:
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT id, booking_month, team_member, capacity_profile_key,
                   system_capacity_hours, leave_hours, availability_hours, booking_hours,
                   notes, created_at_utc, updated_at_utc
            FROM support_hour_booking_headers
            WHERE booking_month = ?
            ORDER BY lower(team_member) ASC
            """,
            (month,),
        ).fetchall()
    finally:
        conn.close()
    return [_row_to_header(row) for row in rows]


def upsert_booking_header(db_path: Path, booking_month: str, team_member: str, payload: dict) -> dict:
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    member = normalize_team_member(team_member)
    raw = payload or {}

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        existing = conn.execute(
            """
            SELECT id, capacity_profile_key, system_capacity_hours, leave_hours, availability_hours, booking_hours, notes
            FROM support_hour_booking_headers
            WHERE booking_month = ? AND team_member = ?
            """,
            (month, member),
        ).fetchone()

        capacity_profile_key = _to_text(raw.get("capacity_profile_key")) or (
            _to_text(existing["capacity_profile_key"]) if existing else ""
        )
        system_capacity_hours = (
            compute_person_month_capacity_hours(db_path, month, capacity_profile_key)
            if capacity_profile_key
            else (_to_float(existing["system_capacity_hours"]) if existing else 0.0)
        )
        leave_hours = _to_float(
            raw.get("leave_hours"), _to_float(existing["leave_hours"]) if existing else DEFAULT_ASSUMED_LEAVE_HOURS
        )
        if leave_hours < 0:
            raise ValueError("leave_hours must be >= 0.")
        computed_availability = round(max(system_capacity_hours - leave_hours, 0.0), 2)
        availability_hours = (
            round(_to_float(raw.get("availability_hours")), 2)
            if raw.get("availability_hours") not in (None, "")
            else computed_availability
        )
        if "booking_hours" in raw and raw.get("booking_hours") not in (None, ""):
            booking_hours = _to_float(raw.get("booking_hours"))
        elif existing:
            booking_hours = _to_float(existing["booking_hours"])
        else:
            booking_hours = availability_hours
        if booking_hours < 0:
            raise ValueError("booking_hours must be >= 0.")
        notes = _to_text(raw.get("notes", existing["notes"] if existing else ""))
        now = _utc_now_iso()

        if existing:
            conn.execute(
                """
                UPDATE support_hour_booking_headers
                SET capacity_profile_key = ?, system_capacity_hours = ?, leave_hours = ?,
                    availability_hours = ?, booking_hours = ?, notes = ?, updated_at_utc = ?
                WHERE booking_month = ? AND team_member = ?
                """,
                (
                    capacity_profile_key, system_capacity_hours, leave_hours,
                    availability_hours, booking_hours, notes, now,
                    month, member,
                ),
            )
        else:
            conn.execute(
                """
                INSERT INTO support_hour_booking_headers (
                    booking_month, team_member, capacity_profile_key,
                    system_capacity_hours, leave_hours, availability_hours, booking_hours,
                    notes, created_at_utc, updated_at_utc
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    month, member, capacity_profile_key,
                    system_capacity_hours, leave_hours, availability_hours, booking_hours,
                    notes, now, now,
                ),
            )
        conn.commit()

        out_row = conn.execute(
            """
            SELECT id, booking_month, team_member, capacity_profile_key,
                   system_capacity_hours, leave_hours, availability_hours, booking_hours,
                   notes, created_at_utc, updated_at_utc
            FROM support_hour_booking_headers
            WHERE booking_month = ? AND team_member = ?
            """,
            (month, member),
        ).fetchone()
    finally:
        conn.close()
    return _row_to_header(out_row)


def delete_booking_header(db_path: Path, booking_month: str, team_member: str) -> bool:
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    member = normalize_team_member(team_member)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            "DELETE FROM support_hour_booking_headers WHERE booking_month = ? AND team_member = ?",
            (month, member),
        )
        conn.execute(
            "DELETE FROM support_hour_booking_allocations WHERE booking_month = ? AND team_member = ?",
            (month, member),
        )
        conn.commit()
    finally:
        conn.close()
    return cur.rowcount > 0


def list_month_allocations(db_path: Path, booking_month: str) -> list[dict]:
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT team_member, project_key, percentage
            FROM support_hour_booking_allocations
            WHERE booking_month = ?
            """,
            (month,),
        ).fetchall()
    finally:
        conn.close()
    return [
        {
            "team_member": _to_text(row["team_member"]),
            "project_key": _to_text(row["project_key"]),
            "percentage": _to_float(row["percentage"]),
        }
        for row in rows
    ]


def upsert_allocation(db_path: Path, booking_month: str, team_member: str, project_key: str, percentage: object) -> dict:
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    member = normalize_team_member(team_member)
    project = normalize_project_key(project_key)
    pct = normalize_percentage(percentage)
    now = _utc_now_iso()

    conn = sqlite3.connect(db_path)
    try:
        if pct == 0:
            conn.execute(
                "DELETE FROM support_hour_booking_allocations WHERE booking_month = ? AND team_member = ? AND project_key = ?",
                (month, member, project),
            )
        else:
            conn.execute(
                """
                INSERT INTO support_hour_booking_allocations (booking_month, team_member, project_key, percentage, updated_at_utc)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(booking_month, team_member, project_key)
                DO UPDATE SET percentage = excluded.percentage, updated_at_utc = excluded.updated_at_utc
                """,
                (month, member, project, pct, now),
            )
        conn.commit()
    finally:
        conn.close()
    return {"booking_month": month, "team_member": member, "project_key": project, "percentage": pct}


def get_month_matrix(db_path: Path, booking_month: str, project_keys: list[str] | None = None) -> dict:
    """Full payload for the UI: headers + allocation % + computed hours-per-project mirror matrix."""
    init_support_booking_db(db_path)
    month = normalize_booking_month(booking_month)
    headers = list_month_headers(db_path, month)
    allocations = list_month_allocations(db_path, month)

    alloc_by_member: dict[str, dict[str, float]] = {}
    for row in allocations:
        alloc_by_member.setdefault(row["team_member"], {})[row["project_key"]] = row["percentage"]

    columns = list(dict.fromkeys(project_keys or []))
    seen_columns = set(columns)
    # make sure any project already allocated (even if since made inactive) still shows up
    for row in allocations:
        project_key = row["project_key"]
        if project_key in seen_columns:
            continue
        seen_columns.add(project_key)
        columns.append(project_key)
    if len(columns) > 1:
        active_prefix = list(dict.fromkeys(project_keys or []))
        active_prefix_set = set(active_prefix)
        extra_columns = [key for key in columns if key not in active_prefix_set]
        extra_columns.sort(key=str.casefold)
        columns = active_prefix + extra_columns

    members_out = []
    for header in headers:
        member = header["team_member"]
        pct_map = alloc_by_member.get(member, {})
        pct_sum = round(sum(pct_map.values()), 4)
        hours_map = {
            project_key: round(header["booking_hours"] * pct_map.get(project_key, 0.0), 2)
            for project_key in columns
        }
        members_out.append(
            {
                **header,
                "allocations": {project_key: pct_map.get(project_key, 0.0) for project_key in columns},
                "hours": hours_map,
                "allocation_pct_total": pct_sum,
                "over_allocated": pct_sum > 1.0001,
                "over_capacity": header["booking_hours"] > header["system_capacity_hours"] + 0.01,
            }
        )

    return {
        "booking_month": month,
        "project_columns": columns,
        "members": members_out,
    }
