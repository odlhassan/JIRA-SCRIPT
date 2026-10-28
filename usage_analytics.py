"""EPR usage analytics: page views and EPR worklog submissions in a standalone SQLite file.

The file is separate from the reporting DBs on purpose; it can be deleted without affecting any report.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
import json
from pathlib import Path
from contextlib import closing
import sqlite3
import threading
from urllib.parse import urlsplit

from flask import g, jsonify, request, send_file

USAGE_DB_FILE = "usage_analytics.db"
RETENTION_DAYS = 400
MY_DAY_PATH = "/my_day.html"
WORKLOG_SUBMIT_PATH = "/api/employee-capacity-utilization/worklog/submit"
SKIP_PATHS = {"/", "/login", "/access-pending"}
_write_lock = threading.Lock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS usage_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    occurred_at_utc TEXT NOT NULL,
    event_date TEXT NOT NULL,
    account_id TEXT NOT NULL DEFAULT '',
    display_name TEXT NOT NULL DEFAULT '',
    role TEXT NOT NULL DEFAULT '',
    event_type TEXT NOT NULL,
    path TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_usage_events_date ON usage_events(event_date);
CREATE INDEX IF NOT EXISTS idx_usage_events_user_date ON usage_events(account_id, display_name, event_date);
"""

# Distinct-person key: Atlassian account when signed in, else the display name (local servers without sign-in).
PERSON = "CASE WHEN account_id <> '' THEN account_id ELSE display_name END"


def init_usage_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(db_path)) as conn, conn:
        conn.executescript(SCHEMA)
        cutoff = (date.today() - timedelta(days=RETENTION_DAYS)).isoformat()
        conn.execute("DELETE FROM usage_events WHERE event_date < ?", (cutoff,))


def normalize_path(path: str) -> str:
    path = "/" + str(path or "").split("?")[0].split("#")[0].lstrip("/")
    if path.startswith("/report_html/") and path.endswith(".html"):
        path = "/" + path.rsplit("/", 1)[-1]
    return path.rstrip("/") or "/"


def record_event(db_path: Path, event_type: str, path: str, user: dict | None, when: datetime | None = None) -> None:
    when = when or datetime.now(timezone.utc)
    user = user or {}
    row = (when.astimezone(timezone.utc).isoformat(), when.astimezone().date().isoformat(),
           str(user.get("account_id") or ""), str(user.get("display_name") or "Local (no sign-in)"),
           str(user.get("role") or ""), event_type, normalize_path(path))
    with _write_lock, closing(sqlite3.connect(db_path)) as conn, conn:
        conn.execute("INSERT INTO usage_events (occurred_at_utc, event_date, account_id, display_name, role, event_type, path)"
                     " VALUES (?, ?, ?, ?, ?, ?, ?)", row)


def _workdays_between(last: date, today: date) -> int:
    count, day = 0, last + timedelta(days=1)
    while day <= today:
        count += day.weekday() < 5
        day += timedelta(days=1)
    return count


def build_usage_summary(db_path: Path, today: date, days: int, teams: list[dict], users: list[dict]) -> dict:
    """Aggregate adoption metrics for the admin screen; users are EPR access records, teams are performance teams."""
    start = today - timedelta(days=days - 1)
    prev_start, prev_end = start - timedelta(days=days), start - timedelta(days=1)
    week_start = today - timedelta(days=6)
    iso = date.isoformat
    with closing(sqlite3.connect(db_path)) as conn, conn:
        def scalar(sql, *args):
            return conn.execute(sql, args).fetchone()[0] or 0

        def actives(a, b):
            return scalar(f"SELECT count(DISTINCT {PERSON}) FROM usage_events WHERE event_type = 'page_view' AND event_date BETWEEN ? AND ?", iso(a), iso(b))

        daily = dict(conn.execute(
            f"SELECT event_date, count(DISTINCT {PERSON}) FROM usage_events WHERE event_type = 'page_view'"
            " AND event_date BETWEEN ? AND ? GROUP BY event_date", (iso(prev_start), iso(today))).fetchall())
        pages = conn.execute(
            f"SELECT path, count(*), count(DISTINCT {PERSON}) FROM usage_events WHERE event_type = 'page_view'"
            " AND event_date BETWEEN ? AND ? GROUP BY path ORDER BY 2 DESC LIMIT 8", (iso(start), iso(today))).fetchall()
        worklogs = scalar("SELECT count(*) FROM usage_events WHERE event_type = 'worklog_submitted' AND event_date BETWEEN ? AND ?", iso(start), iso(today))
        prev_worklogs = scalar("SELECT count(*) FROM usage_events WHERE event_type = 'worklog_submitted' AND event_date BETWEEN ? AND ?", iso(prev_start), iso(prev_end))
        my_day_worklogs = scalar("SELECT count(*) FROM usage_events WHERE event_type = 'worklog_submitted' AND path = ? AND event_date BETWEEN ? AND ?", MY_DAY_PATH, iso(start), iso(today))
        period_users = actives(start, today)
        prev_users = actives(prev_start, prev_end)
        my_day_users = scalar(f"SELECT count(DISTINCT {PERSON}) FROM usage_events WHERE event_type = 'page_view' AND path = ? AND event_date BETWEEN ? AND ?", MY_DAY_PATH, iso(start), iso(today))
        prev_my_day_users = scalar(f"SELECT count(DISTINCT {PERSON}) FROM usage_events WHERE event_type = 'page_view' AND path = ? AND event_date BETWEEN ? AND ?", MY_DAY_PATH, iso(prev_start), iso(prev_end))
        launch = scalar("SELECT min(event_date) FROM usage_events WHERE path = ?", MY_DAY_PATH) or ""
        wau, prev_wau = actives(week_start, today), actives(week_start - timedelta(days=7), week_start - timedelta(days=1))
        active_days = conn.execute(
            "SELECT account_id, lower(trim(display_name)), count(DISTINCT event_date), max(event_date) FROM usage_events"
            " WHERE event_type = 'page_view' AND event_date >= ? GROUP BY account_id, lower(trim(display_name))",
            (iso(today - timedelta(days=RETENTION_DAYS)),)).fetchall()
        recent_days = conn.execute(
            "SELECT lower(trim(display_name)), count(DISTINCT event_date) FROM usage_events"
            " WHERE event_type = 'page_view' AND event_date BETWEEN ? AND ? GROUP BY 1", (iso(week_start), iso(today))).fetchall()

    def avg_dau(a: date, b: date) -> float:
        values = [daily.get(iso(a + timedelta(days=i)), 0) for i in range((b - a).days + 1) if (a + timedelta(days=i)).weekday() < 5]
        return round(sum(values) / len(values), 1) if values else 0.0

    week_avg, prev_week_avg = avg_dau(week_start, today), avg_dau(week_start - timedelta(days=7), week_start - timedelta(days=1))
    stickiness = round(week_avg / wau * 100) if wau else 0
    prev_stickiness = round(prev_week_avg / prev_wau * 100) if prev_wau else 0
    my_day_share = round(my_day_users / period_users * 100) if period_users else 0
    prev_my_day_share = round(prev_my_day_users / prev_users * 100) if prev_users else 0

    series = []
    for i in range(days):
        day = start + timedelta(days=i)
        series.append({"date": iso(day), "users": daily.get(iso(day), 0), "weekend": day.weekday() >= 5, "launch": iso(day) == launch})

    by_account = {row[0]: row for row in active_days if row[0]}
    by_name = {row[1]: row for row in active_days}
    week_days_by_name = dict(recent_days)
    team_rows = []
    for team in teams:
        members = [m for m in team.get("assignees") or [] if str(m).strip()]
        if not members:
            continue
        engaged = sum(1 for m in members if week_days_by_name.get(str(m).strip().casefold(), 0) >= 3)
        team_rows.append({"team": team.get("team_name", ""), "members": len(members), "engaged": engaged,
                          "pct": round(engaged / len(members) * 100)})
    team_rows.sort(key=lambda t: t["pct"], reverse=True)

    quiet = []
    for user in users:
        if not user.get("approved") or not user.get("active"):
            continue
        row = by_account.get(user.get("account_id")) or by_name.get(str(user.get("display_name") or "").strip().casefold())
        last = row[3] if row else ""
        idle = _workdays_between(date.fromisoformat(last), today) if last else None
        if idle is None or idle >= 5:
            quiet.append({"display_name": user.get("display_name", ""), "role": user.get("role", ""), "last_seen": last, "idle_workdays": idle})
    quiet.sort(key=lambda q: (q["idle_workdays"] is not None, -(q["idle_workdays"] or 0)))

    return {
        "range_days": days,
        "from": iso(start),
        "to": iso(today),
        "kpis": {
            "daily_active": {"value": daily.get(iso(today), 0), "avg": week_avg, "previous": prev_week_avg},
            "weekly_active": {"value": wau, "previous": prev_wau},
            "stickiness": {"value": stickiness, "previous": prev_stickiness},
            "my_day_reach": {"value": my_day_share, "previous": prev_my_day_share, "users": my_day_users},
            "worklogs": {"value": worklogs, "previous": prev_worklogs, "via_my_day": my_day_worklogs},
        },
        "period_users": period_users,
        "series": series,
        "my_day_launch": launch,
        "pages": [{"path": p, "views": v, "users": u} for p, v, u in pages],
        "teams": team_rows,
        "quiet": quiet[:12],
        "quiet_total": len(quiet),
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }


def register_usage_analytics(app, db_path: Path, base_dir: Path, list_teams) -> None:
    init_usage_db(db_path)
    app.config["EPR_USAGE_DB"] = db_path

    @app.after_request
    def track_usage(response):
        try:
            path = request.path
            user = getattr(g, "epr_user", None)
            if app.config.get("EPR_AUTH_ENABLED") and not user:
                return response
            if (request.method == "GET" and response.status_code == 200 and response.mimetype == "text/html"
                    and request.headers.get("Sec-Fetch-Dest", "document") in ("document", "")
                    and normalize_path(path) not in SKIP_PATHS and not path.startswith("/api/")):
                record_event(db_path, "page_view", path, user)
            elif request.method == "POST" and path == WORKLOG_SUBMIT_PATH and response.status_code == 200:
                if (response.get_json(silent=True) or {}).get("worklog_id"):
                    record_event(db_path, "worklog_submitted", urlsplit(request.headers.get("Referer", "")).path, user)
        except (sqlite3.Error, OSError, ValueError):
            pass  # Analytics must never break a page or a Jira submission.
        return response

    @app.get("/settings/usage-analytics")
    def usage_analytics_page():
        return send_file(base_dir / "usage_analytics.html")

    @app.get("/api/admin/usage-analytics")
    def usage_analytics_api():
        try:
            days = int(request.args.get("days", 30))
        except ValueError:
            days = 30
        days = days if days in (7, 30, 90) else 30
        users_file = app.config.get("EPR_AUTH_DIRECTORY")
        users_file = Path(users_file) / "users.json" if users_file else None
        users = list(json.loads(users_file.read_text(encoding="utf-8")).values()) if users_file and users_file.exists() else []
        try:
            teams = list_teams()
        except sqlite3.Error:
            teams = []
        return jsonify(build_usage_summary(db_path, date.today(), days, teams, users))
