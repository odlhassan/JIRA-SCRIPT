from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from support_booking_registry import (
    compute_person_month_capacity_hours,
    delete_booking_header,
    get_month_matrix,
    init_month_bookings,
    init_support_booking_db,
    list_capacity_profile_options,
    list_month_allocations,
    list_month_headers,
    normalize_booking_month,
    normalize_percentage,
    upsert_allocation,
    upsert_booking_header,
)


def _seed_capacity_profile(db: Path, from_date: str, to_date: str, hours_per_day: float = 8.0) -> None:
    import sqlite3
    from datetime import datetime, timezone

    conn = sqlite3.connect(db)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS assignee_capacity_settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                from_date TEXT NOT NULL,
                to_date TEXT NOT NULL,
                employee_count INTEGER NOT NULL,
                standard_hours_per_day REAL NOT NULL,
                ramadan_start_date TEXT,
                ramadan_end_date TEXT,
                ramadan_hours_per_day REAL NOT NULL,
                holiday_dates_json TEXT NOT NULL,
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        now = datetime.now(timezone.utc).isoformat()
        conn.execute(
            """
            INSERT INTO assignee_capacity_settings (
                from_date, to_date, employee_count, standard_hours_per_day,
                ramadan_start_date, ramadan_end_date, ramadan_hours_per_day,
                holiday_dates_json, created_at_utc, updated_at_utc
            ) VALUES (?, ?, ?, ?, NULL, NULL, 6.5, '[]', ?, ?)
            """,
            (from_date, to_date, 8, hours_per_day, now, now),
        )
        conn.commit()
    finally:
        conn.close()


def _seed_support_team(db: Path, members: list[str]) -> None:
    import json
    import sqlite3
    from datetime import datetime, timezone

    conn = sqlite3.connect(db)
    try:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS support_team_config (key TEXT, members_json TEXT, updated_at TEXT)"
        )
        conn.execute(
            "INSERT INTO support_team_config (key, members_json, updated_at) VALUES ('members', ?, ?)",
            (json.dumps(members), datetime.now(timezone.utc).isoformat()),
        )
        conn.commit()
    finally:
        conn.close()


class SupportBookingRegistryTests(unittest.TestCase):
    def test_normalize_booking_month(self):
        self.assertEqual(normalize_booking_month("2026-07"), "2026-07")
        with self.assertRaises(ValueError):
            normalize_booking_month("2026-7")
        with self.assertRaises(ValueError):
            normalize_booking_month("bad")

    def test_normalize_percentage_bounds(self):
        self.assertEqual(normalize_percentage(0.3), 0.3)
        with self.assertRaises(ValueError):
            normalize_percentage(-0.1)
        with self.assertRaises(ValueError):
            normalize_percentage(30)  # whole-percent typo guard

    def test_compute_person_month_capacity_hours_july_23_workdays(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31", hours_per_day=8.0)
            hours = compute_person_month_capacity_hours(db, "2026-07", "2026-01-01|2026-12-31")
            # July 2026 has 23 weekdays (Wed 1 Jul .. Fri 31 Jul, no holidays seeded).
            self.assertEqual(hours, 23 * 8.0)

    def test_init_month_bookings_creates_header_per_member_with_default_leave(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            _seed_support_team(db, ["Ameer Hamza", "Abbas"])

            headers = init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")
            self.assertEqual(len(headers), 2)
            ameer = next(h for h in headers if h["team_member"] == "Ameer Hamza")
            self.assertEqual(ameer["system_capacity_hours"], 184.0)  # 23 workdays x 8h
            self.assertEqual(ameer["leave_hours"], 16.0)
            self.assertEqual(ameer["availability_hours"], 168.0)
            self.assertEqual(ameer["booking_hours"], 168.0)

            # Calling init again should not duplicate or reset manual edits.
            upsert_booking_header(db, "2026-07", "Ameer Hamza", {"booking_hours": 168.0})
            headers_again = init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")
            self.assertEqual(len(headers_again), 2)

    def test_upsert_booking_header_manual_override(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            _seed_support_team(db, ["Abbas"])
            init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")

            # Admin manually types 84 hours because Abbas is half on support this month.
            updated = upsert_booking_header(db, "2026-07", "Abbas", {"booking_hours": 84})
            self.assertEqual(updated["booking_hours"], 84.0)
            self.assertEqual(updated["system_capacity_hours"], 184.0)

    def test_allocations_and_mirror_matrix(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            _seed_support_team(db, ["Abbas"])
            init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")
            upsert_booking_header(db, "2026-07", "Abbas", {"booking_hours": 84})

            upsert_allocation(db, "2026-07", "Abbas", "OmniConnect", 0.3)
            upsert_allocation(db, "2026-07", "Abbas", "OmniChat", 0.3)
            upsert_allocation(db, "2026-07", "Abbas", "FintechFuel", 0.2)
            upsert_allocation(db, "2026-07", "Abbas", "DigitalLog", 0.2)

            matrix = get_month_matrix(db, "2026-07", project_keys=["OMNICONNECT", "OMNICHAT", "FINTECHFUEL", "DIGITALLOG"])
            abbas = next(m for m in matrix["members"] if m["team_member"] == "Abbas")
            self.assertEqual(abbas["hours"]["OMNICONNECT"], 25.2)
            self.assertEqual(abbas["hours"]["OMNICHAT"], 25.2)
            self.assertEqual(abbas["hours"]["FINTECHFUEL"], 16.8)
            self.assertEqual(abbas["hours"]["DIGITALLOG"], 16.8)
            self.assertEqual(abbas["allocation_pct_total"], 1.0)
            self.assertFalse(abbas["over_allocated"])

    def test_allocation_zero_percentage_removes_row(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_support_booking_db(db)
            upsert_allocation(db, "2026-07", "Abbas", "OMNICONNECT", 0.5)
            self.assertEqual(len(list_month_allocations(db, "2026-07")), 1)
            upsert_allocation(db, "2026-07", "Abbas", "OMNICONNECT", 0)
            self.assertEqual(len(list_month_allocations(db, "2026-07")), 0)

    def test_over_allocation_flagged(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            _seed_support_team(db, ["Abbas"])
            init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")
            upsert_allocation(db, "2026-07", "Abbas", "OMNICONNECT", 0.7)
            upsert_allocation(db, "2026-07", "Abbas", "OMNICHAT", 0.5)
            matrix = get_month_matrix(db, "2026-07", project_keys=["OMNICONNECT", "OMNICHAT"])
            abbas = next(m for m in matrix["members"] if m["team_member"] == "Abbas")
            self.assertTrue(abbas["over_allocated"])

    def test_delete_booking_header_removes_allocations_too(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            _seed_support_team(db, ["Abbas"])
            init_month_bookings(db, "2026-07", "2026-01-01|2026-12-31")
            upsert_allocation(db, "2026-07", "Abbas", "OMNICONNECT", 0.5)

            deleted = delete_booking_header(db, "2026-07", "Abbas")
            self.assertTrue(deleted)
            self.assertEqual(list_month_headers(db, "2026-07"), [])
            self.assertEqual(list_month_allocations(db, "2026-07"), [])

    def test_list_capacity_profile_options(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _seed_capacity_profile(db, "2026-01-01", "2026-12-31")
            options = list_capacity_profile_options(db)
            self.assertEqual(len(options), 1)
            self.assertEqual(options[0]["capacity_profile_key"], "2026-01-01|2026-12-31")


if __name__ == "__main__":
    unittest.main()
