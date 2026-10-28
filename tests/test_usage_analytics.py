from datetime import date, datetime, timezone
import json
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

from flask import Flask, jsonify

from epr_auth import OWNER_ACCOUNT_ID, register_epr_auth
from jira_oauth import register_jira_oauth
from usage_analytics import build_usage_summary, init_usage_db, normalize_path, record_event, register_usage_analytics


def at(day: str) -> datetime:
    return datetime.fromisoformat(day + "T09:00:00").astimezone(timezone.utc)


class UsageSummaryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.tmp.cleanup)
        self.db = Path(self.tmp.name) / "usage_analytics.db"
        init_usage_db(self.db)

    def test_normalize_path_collapses_report_html_copies(self):
        self.assertEqual(normalize_path("/report_html/my_day.html?x=1"), "/my_day.html")
        self.assertEqual(normalize_path("/settings/usage-analytics/"), "/settings/usage-analytics")

    def test_summary_counts_people_pages_teams_and_quiet_users(self):
        ayesha = {"account_id": "a1", "display_name": "Ayesha Khan", "role": "member"}
        bilal = {"account_id": "b1", "display_name": "Bilal Ahmed", "role": "member"}
        for day in ("2026-10-26", "2026-10-27", "2026-10-28"):
            record_event(self.db, "page_view", "/my_day.html", ayesha, at(day))
        record_event(self.db, "page_view", "/report_html/dashboard.html", ayesha, at("2026-10-28"))
        record_event(self.db, "page_view", "/dashboard.html", bilal, at("2026-10-12"))
        record_event(self.db, "worklog_submitted", "/my_day.html", ayesha, at("2026-10-28"))
        users = [dict(ayesha, approved=True, active=True), dict(bilal, approved=True, active=True),
                 {"account_id": "c1", "display_name": "Chris Never", "approved": True, "active": True},
                 {"account_id": "d1", "display_name": "Pending", "approved": False, "active": True}]
        teams = [{"team_name": "Platform", "assignees": ["Ayesha Khan", "Bilal Ahmed"]}, {"team_name": "Empty", "assignees": []}]

        summary = build_usage_summary(self.db, date(2026, 10, 28), 30, teams, users)

        self.assertEqual(summary["kpis"]["daily_active"]["value"], 1)
        self.assertEqual(summary["kpis"]["weekly_active"]["value"], 1)
        self.assertEqual(summary["period_users"], 2)
        self.assertEqual(summary["kpis"]["my_day_reach"]["value"], 50)
        self.assertEqual(summary["kpis"]["worklogs"], {"value": 1, "previous": 0, "via_my_day": 1})
        self.assertEqual(summary["my_day_launch"], "2026-10-26")
        self.assertEqual(summary["pages"][0], {"path": "/my_day.html", "views": 3, "users": 1})
        self.assertEqual(summary["pages"][1]["path"], "/dashboard.html")
        self.assertEqual(summary["teams"], [{"team": "Platform", "members": 2, "engaged": 1, "pct": 50}])
        self.assertEqual([q["display_name"] for q in summary["quiet"]], ["Chris Never", "Bilal Ahmed"])
        self.assertEqual(summary["quiet"][1]["idle_workdays"], 12)
        self.assertEqual(len(summary["series"]), 30)
        self.assertTrue(any(point["launch"] for point in summary["series"]))


class UsageTrackingRouteTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.oauth = root / "oauth"
        self.oauth.mkdir()
        env = patch.dict(os.environ, {"EPR_AUTH_ENABLED": "1", "JIRA_OAUTH_CLIENT_ID": "", "JIRA_OAUTH_CLIENT_SECRET": "", "JIRA_OAUTH_REDIRECT_URI": ""})
        env.start()
        self.addCleanup(env.stop)
        (root / "usage_analytics.html").write_text("<!doctype html><title>Usage</title>", encoding="utf-8")
        self.app = Flask(__name__)
        self.app.testing = True
        register_jira_oauth(self.app, self.oauth)
        register_epr_auth(self.app, self.oauth, {"my_day.html"})
        self.db = root / "usage_analytics.db"
        register_usage_analytics(self.app, self.db, root, lambda: [])
        self.app.add_url_rule("/my_day.html", "my_day", lambda: "<!doctype html><p>My Day</p>")
        self.app.add_url_rule("/api/employee-capacity-utilization/worklog/submit", "submit",
                              lambda: jsonify(worklog_id="10001"), methods=["POST"])
        self.client = self.app.test_client()

    def login(self, account, name):
        sid = ("a" if account == OWNER_ACCOUNT_ID else "b") * 64
        (self.oauth / (sid + ".json")).write_text(json.dumps({
            "session_expires_at": time.time() + 86400, "epr_signed_at": time.time(), "epr_csrf": "csrf",
            "epr_identity": {"account_id": account, "display_name": name, "avatar_url": ""}}), encoding="utf-8")
        self.client.set_cookie("epr_jira_session", sid)

    def rows(self):
        import sqlite3
        from contextlib import closing
        with closing(sqlite3.connect(self.db)) as conn:
            return conn.execute("SELECT event_type, path, display_name FROM usage_events ORDER BY id").fetchall()

    def test_page_views_and_worklogs_are_recorded_for_signed_in_users_only(self):
        self.client.get("/my_day.html")
        self.assertEqual(self.rows(), [])
        self.login(OWNER_ACCOUNT_ID, "Hassan Malik")
        self.assertEqual(self.client.get("/my_day.html", headers={"Sec-Fetch-Dest": "document"}).status_code, 200)
        self.client.get("/my_day.html", headers={"Sec-Fetch-Dest": "iframe"})
        self.client.post("/api/employee-capacity-utilization/worklog/submit", json={},
                         headers={"Origin": "http://localhost", "Referer": "http://localhost/my_day.html"})
        self.assertEqual(self.rows(), [("page_view", "/my_day.html", "Hassan Malik"),
                                       ("worklog_submitted", "/my_day.html", "Hassan Malik")])

    def test_admin_screen_and_api_are_admin_only(self):
        self.login(OWNER_ACCOUNT_ID, "Hassan Malik")
        self.assertEqual(self.client.get("/settings/usage-analytics").status_code, 200)
        api = self.client.get("/api/admin/usage-analytics?days=7")
        self.assertEqual(api.status_code, 200)
        self.assertEqual(api.json["range_days"], 7)
        self.assertEqual(api.json["quiet"], [])
        users_file = self.oauth.parent / "epr_access" / "users.json"
        users = json.loads(users_file.read_text(encoding="utf-8"))
        users["member-1"] = {"account_id": "member-1", "display_name": "Member One", "role": "member", "approved": True, "active": True}
        users_file.write_text(json.dumps(users), encoding="utf-8")
        self.login("member-1", "Member One")
        self.assertEqual(self.client.get("/settings/usage-analytics").status_code, 403)
        self.assertEqual(self.client.get("/api/admin/usage-analytics").status_code, 403)


if __name__ == "__main__":
    unittest.main()
