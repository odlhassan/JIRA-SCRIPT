from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import report_server
from report_server import _init_epics_management_db, create_report_server_app


class ProductReleasesApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["FF", "2026-05-01", "2026-05-01", "2026-W18", "2026-05", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def _seed_project(self, client, project_key: str = "FF", project_name: str = "Fintech Fuel"):
        resp = client.post(
            "/api/projects",
            json={
                "project_key": project_key,
                "project_name": project_name,
                "display_name": project_name,
                "color_hex": "#1D4ED8",
            },
        )
        self.assertIn(resp.status_code, (200, 409))

    def _seed_epic(self, db_path: Path, row_id: str = "19", epic_key: str = "FF-274", project_key: str = "FF"):
        _init_epics_management_db(db_path)
        conn = sqlite3.connect(db_path)
        try:
            conn.execute(
                "INSERT INTO epics_management (id, epic_key, project_key, project_name, product_category, epic_name) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (row_id, epic_key, project_key, "Fintech Fuel", "Payments", "Site Diagnostics Dashboard"),
            )
            conn.commit()
        finally:
            conn.close()

    def test_release_action_updates_release_status_and_logs_actor(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            self._seed_project(client)
            self._seed_epic(root / "assignee_hours_capacity.db")

            create_resp = client.post(
                "/api/product-releases",
                json={
                    "project_key": "FF",
                    "release_number": "v9",
                    "release_date": "2026-05-22",
                    "notes": "Initial bucket",
                },
            )
            self.assertEqual(create_resp.status_code, 201)
            release_id = create_resp.get_json()["release"]["id"]

            assign_resp = client.post(
                f"/api/product-releases/{release_id}/epics",
                json={"epic_row_id": "19", "epic_type": "new_feature"},
            )
            self.assertEqual(assign_resp.status_code, 201)

            action_resp = client.post(
                f"/api/product-releases/{release_id}/actions",
                json={
                    "action": "Released",
                    "actual_date": "2026-05-24",
                    "actor": "Hassan",
                    "notes": "Released to production",
                },
            )
            self.assertEqual(action_resp.status_code, 200)
            release = action_resp.get_json()
            self.assertEqual(release["release_status"], "released")
            self.assertEqual(len(release["actions"]), 1)
            self.assertEqual(release["actions"][0]["actor"], "Hassan")
            self.assertEqual(release["actions"][0]["action"], "released")
            self.assertEqual(release["actions"][0]["actual_date"], "2026-05-24")

            conn = sqlite3.connect(root / "assignee_hours_capacity.db")
            try:
                row = conn.execute(
                    "SELECT release_status, actual_release_date FROM product_releases WHERE id = ?",
                    (release_id,),
                ).fetchone()
            finally:
                conn.close()
            self.assertEqual(row[0], "released")
            self.assertEqual(row[1], "2026-05-24")

    def test_reschedule_and_shelved_actions_are_logged_on_release(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            self._seed_project(client)
            self._seed_epic(root / "assignee_hours_capacity.db")

            create_resp = client.post(
                "/api/product-releases",
                json={
                    "project_key": "FF",
                    "release_number": "v10",
                    "release_date": "2026-06-18",
                    "notes": "",
                },
            )
            release_id = create_resp.get_json()["release"]["id"]
            client.post(
                f"/api/product-releases/{release_id}/epics",
                json={"epic_row_id": "19", "epic_type": "enhancement"},
            )

            reschedule_resp = client.post(
                f"/api/product-releases/{release_id}/actions",
                json={
                    "action": "Rescheduled",
                    "actual_date": "2026-06-25",
                    "actor": "Hassan",
                    "notes": "QA found blockers",
                },
            )
            self.assertEqual(reschedule_resp.status_code, 200)
            rescheduled = reschedule_resp.get_json()
            self.assertEqual(rescheduled["release_status"], "scheduled")
            self.assertEqual(rescheduled["actions"][0]["action"], "rescheduled")
            self.assertEqual(rescheduled["actions"][0]["actual_date"], "2026-06-25")
            self.assertEqual(rescheduled["release_date"], "2026-06-25")
            self.assertEqual(rescheduled["previous_release_date"], "2026-06-18")

            shelf_resp = client.post(
                f"/api/product-releases/{release_id}/actions",
                json={
                    "action": "shelved",
                    "actor": "Hassan",
                    "notes": "Scope moved to later planning window",
                },
            )
            self.assertEqual(shelf_resp.status_code, 200)
            shelfed = shelf_resp.get_json()
            self.assertEqual(shelfed["release_status"], "shelved")
            self.assertEqual(len(shelfed["actions"]), 2)
            self.assertEqual(shelfed["actions"][0]["action"], "shelved")

            pool_resp = client.get("/api/product-releases/epics/pool")
            self.assertEqual(pool_resp.status_code, 200)
            epics = pool_resp.get_json()["epics"]
            self.assertEqual(epics[0]["release_id"], release_id)

            conn = sqlite3.connect(root / "assignee_hours_capacity.db")
            try:
                row = conn.execute(
                    "SELECT release_status, release_date, previous_release_date FROM product_releases WHERE id = ?",
                    (release_id,),
                ).fetchone()
            finally:
                conn.close()
            self.assertEqual(row[0], "shelved")
            self.assertEqual(row[1], "2026-06-25")
            self.assertEqual(row[2], "2026-06-18")


if __name__ == "__main__":
    unittest.main()
