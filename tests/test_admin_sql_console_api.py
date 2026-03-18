from __future__ import annotations

import sqlite3
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

from report_server import SQL_CONSOLE_MAX_ROWS, create_report_server_app


class AdminSqlConsoleApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_schema_api_lists_tables_for_canonical_and_exports(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            with sqlite3.connect(root / "assignee_hours_capacity.db") as conn:
                conn.execute("CREATE TABLE canonical_demo(id INTEGER PRIMARY KEY, name TEXT NOT NULL)")
                conn.commit()
            with sqlite3.connect(root / "jira_exports.db") as conn:
                conn.execute("CREATE TABLE export_demo(issue_key TEXT, hours_logged REAL)")
                conn.commit()

            canonical_resp = client.get("/api/admin/sql-console/schema?database=canonical")
            self.assertEqual(canonical_resp.status_code, 200)
            canonical_body = canonical_resp.get_json() or {}
            canonical_tables = {str(item.get("name")): item for item in canonical_body.get("tables") or []}
            self.assertIn("canonical_demo", canonical_tables)
            canonical_cols = canonical_tables["canonical_demo"].get("columns") or []
            self.assertEqual(canonical_cols[0]["name"], "id")

            exports_resp = client.get("/api/admin/sql-console/schema?database=exports")
            self.assertEqual(exports_resp.status_code, 200)
            exports_body = exports_resp.get_json() or {}
            exports_tables = {str(item.get("name")): item for item in exports_body.get("tables") or []}
            self.assertIn("export_demo", exports_tables)

    def test_execute_allows_read_only_queries_for_both_targets(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            with sqlite3.connect(root / "assignee_hours_capacity.db") as conn:
                conn.execute("CREATE TABLE canonical_demo(name TEXT, hours REAL)")
                conn.execute("INSERT INTO canonical_demo(name, hours) VALUES ('alice', 3.5)")
                conn.commit()
            with sqlite3.connect(root / "jira_exports.db") as conn:
                conn.execute("CREATE TABLE export_demo(issue_key TEXT, status TEXT)")
                conn.execute("INSERT INTO export_demo(issue_key, status) VALUES ('O2-1', 'Open')")
                conn.commit()

            canonical_resp = client.post(
                "/api/admin/sql-console/execute",
                json={"database": "canonical", "sql": "SELECT name, hours FROM canonical_demo"},
            )
            self.assertEqual(canonical_resp.status_code, 200)
            canonical_body = canonical_resp.get_json() or {}
            self.assertTrue(canonical_body.get("ok"))
            self.assertEqual(canonical_body.get("columns"), ["name", "hours"])
            self.assertEqual((canonical_body.get("rows") or [])[0]["name"], "alice")

            exports_resp = client.post(
                "/api/admin/sql-console/execute",
                json={"database": "exports", "sql": "SELECT issue_key, status FROM export_demo"},
            )
            self.assertEqual(exports_resp.status_code, 200)
            exports_body = exports_resp.get_json() or {}
            self.assertTrue(exports_body.get("ok"))
            self.assertEqual((exports_body.get("rows") or [])[0]["issue_key"], "O2-1")

    def test_execute_rejects_write_sql_empty_sql_and_multiple_statements(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            with sqlite3.connect(root / "assignee_hours_capacity.db") as conn:
                conn.execute("CREATE TABLE canonical_demo(name TEXT)")
                conn.commit()

            bad_queries = [
                "",
                "UPDATE canonical_demo SET name = 'x'",
                "DELETE FROM canonical_demo",
                "INSERT INTO canonical_demo(name) VALUES ('x')",
                "DROP TABLE canonical_demo",
                "ALTER TABLE canonical_demo ADD COLUMN note TEXT",
                "ATTACH DATABASE 'other.db' AS other",
                "SELECT 1; SELECT 2",
            ]
            for sql in bad_queries:
                with self.subTest(sql=sql):
                    response = client.post(
                        "/api/admin/sql-console/execute",
                        json={"database": "canonical", "sql": sql},
                    )
                    self.assertEqual(response.status_code, 400)
                    body = response.get_json() or {}
                    self.assertFalse(body.get("ok", False))

    def test_execute_truncates_large_result_sets(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            with sqlite3.connect(root / "jira_exports.db") as conn:
                conn.execute("CREATE TABLE export_demo(id INTEGER)")
                conn.executemany("INSERT INTO export_demo(id) VALUES (?)", [(idx,) for idx in range(SQL_CONSOLE_MAX_ROWS + 25)])
                conn.commit()

            response = client.post(
                "/api/admin/sql-console/execute",
                json={"database": "exports", "sql": "SELECT id FROM export_demo ORDER BY id"},
            )
            self.assertEqual(response.status_code, 200)
            body = response.get_json() or {}
            self.assertTrue(body.get("truncated"))
            self.assertEqual(body.get("row_count"), SQL_CONSOLE_MAX_ROWS)
            self.assertEqual(len(body.get("rows") or []), SQL_CONSOLE_MAX_ROWS)

    def test_export_downloads_excel_for_read_only_query(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            with sqlite3.connect(root / "assignee_hours_capacity.db") as conn:
                conn.execute("CREATE TABLE canonical_demo(name TEXT, hours REAL)")
                conn.execute("INSERT INTO canonical_demo(name, hours) VALUES ('alice', 3.5)")
                conn.execute("INSERT INTO canonical_demo(name, hours) VALUES ('bob', 4.0)")
                conn.commit()

            response = client.post(
                "/api/admin/sql-console/export",
                json={"database": "canonical", "sql": "SELECT name, hours FROM canonical_demo ORDER BY name"},
            )
            self.assertEqual(response.status_code, 200)
            self.assertIn(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.headers.get("Content-Type", ""),
            )
            self.assertIn("attachment;", response.headers.get("Content-Disposition", ""))

            workbook = load_workbook(BytesIO(response.data))
            worksheet = workbook.active
            self.assertEqual(worksheet.title, "Query Results")
            self.assertEqual([worksheet["A1"].value, worksheet["B1"].value], ["name", "hours"])
            self.assertEqual([worksheet["A2"].value, worksheet["B2"].value], ["alice", 3.5])
            self.assertEqual([worksheet["A3"].value, worksheet["B3"].value], ["bob", 4.0])

    def test_missing_target_db_returns_404(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            response = client.get("/api/admin/sql-console/schema?database=exports")
            self.assertEqual(response.status_code, 404)
            body = response.get_json() or {}
            self.assertFalse(body.get("ok", False))


if __name__ == "__main__":
    unittest.main()
