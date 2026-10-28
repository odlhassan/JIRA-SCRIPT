"""My Day: one employee's tasks, hours, gaps and upcoming leave, read from the canonical DB (read-only)."""
from __future__ import annotations

from datetime import date, datetime, timedelta
import json
from pathlib import Path
from contextlib import closing
import sqlite3

from canonical_report_data import build_rlt_leave_snapshot, resolve_canonical_run_id

DONE_STATUSES = {"done", "closed", "resolved", "resolved!", "bug fixed", "not a bug"}
STREAK_WEEKS = 26
TASK_LIMIT = 10
# A day short by at most this much still counts as complete (breaks, rounding of 7.5h days).
DAY_GRACE_HOURS = 0.5


def _key(value) -> str:
    return str(value or "").strip().casefold()


def _iso(value: date) -> str:
    return value.isoformat()


def _profile_for(profiles: list[dict], day: date) -> dict:
    iso = _iso(day)
    for profile in profiles:
        if profile["from_date"] <= iso <= profile["to_date"]:
            return profile
    return profiles[0] if profiles else {"standard_hours_per_day": 8.0, "holiday_dates": []}


def _load_profiles(conn: sqlite3.Connection) -> list[dict]:
    try:
        rows = conn.execute(
            "SELECT from_date, to_date, standard_hours_per_day, ramadan_start_date, ramadan_end_date,"
            " ramadan_hours_per_day, holiday_dates_json FROM assignee_capacity_settings ORDER BY updated_at_utc DESC"
        ).fetchall()
    except sqlite3.Error:
        return []
    out = []
    for row in rows:
        try:
            holidays = [str(x) for x in json.loads(row[6] or "[]") if x]
        except (TypeError, ValueError):
            holidays = []
        out.append({"from_date": str(row[0] or ""), "to_date": str(row[1] or "9999-12-31"),
                    "standard_hours_per_day": float(row[2] or 8), "ramadan_start_date": str(row[3] or ""),
                    "ramadan_end_date": str(row[4] or ""), "ramadan_hours_per_day": float(row[5] or 0),
                    "holiday_dates": holidays})
    return out


def expected_hours(profiles: list[dict], day: date) -> tuple[float, str]:
    """Hours an employee is expected to log on a day, and why (workday, weekend, holiday, ramadan)."""
    if day.weekday() >= 5:
        return 0.0, "weekend"
    profile = _profile_for(profiles, day)
    iso = _iso(day)
    if iso in set(profile.get("holiday_dates") or []):
        return 0.0, "holiday"
    start, end = profile.get("ramadan_start_date"), profile.get("ramadan_end_date")
    if start and end and start <= iso <= end:
        return float(profile.get("ramadan_hours_per_day") or 0), "ramadan"
    return float(profile.get("standard_hours_per_day") or 8), "workday"


def day_status(day: date, today: date, data_through: str, expected: float, logged: float, leave: float) -> tuple[str, float]:
    """Classify one calendar day; returns (status, missing_hours)."""
    if day > today:
        return ("leave" if leave > 0 else "future"), 0.0
    if expected <= 0:
        return "off", 0.0
    if leave >= expected - 0.01:
        return "leave", 0.0
    if logged + leave >= expected - DAY_GRACE_HOURS:
        return "complete", 0.0
    if day == today:
        return "today", 0.0
    if _iso(day) > data_through:
        return "unsynced", 0.0
    return "missing", round(expected - logged - leave, 2)


def _monday(day: date) -> date:
    return day - timedelta(days=day.weekday())


def fetch_live_week_worklogs(session, base_url: str, account_id: str, today: date, max_issues: int = 40) -> list[dict]:
    """This week's worklogs by one Jira author, read live so My Day reflects entries made since the last refresh."""
    monday = _monday(today)
    jql = f'worklogAuthor = "{account_id}" AND worklogDate >= "{_iso(monday)}" AND worklogDate <= "{_iso(today)}"'
    response = session.get(base_url + "/rest/api/3/search/jql", params={"jql": jql, "fields": "summary", "maxResults": max_issues},
                           timeout=(10, 25))
    response.raise_for_status()
    out = []
    for issue in response.json().get("issues") or []:
        key = str(issue.get("key") or "")
        if not key or key.startswith("RLT-"):
            continue
        summary = str((issue.get("fields") or {}).get("summary") or "")
        page = session.get(base_url + f"/rest/api/3/issue/{key}/worklog",
                           params={"startedAfter": int(datetime.combine(monday, datetime.min.time()).timestamp() * 1000) - 86400000,
                                   "maxResults": 1000}, timeout=(10, 25))
        page.raise_for_status()
        for row in page.json().get("worklogs") or []:
            if (row.get("author") or {}).get("accountId") != account_id:
                continue
            started = datetime.fromisoformat(str(row.get("started", "")))
            day = started.astimezone().date() if started.tzinfo else started.date()
            if monday <= day <= today:
                out.append({"issue_key": key, "date": _iso(day), "hours": float(row.get("timeSpentSeconds") or 0) / 3600,
                            "summary": summary})
    return out


def list_people(db_path: Path) -> list[str]:
    run_id = resolve_canonical_run_id(db_path)
    if not run_id:
        return []
    with closing(sqlite3.connect(db_path)) as conn, conn:
        rows = conn.execute(
            "SELECT DISTINCT trim(worklog_author) FROM canonical_worklogs WHERE run_id = ? AND trim(worklog_author) <> ''"
            " ORDER BY 1 COLLATE NOCASE", (run_id,)).fetchall()
    return [row[0] for row in rows]


def build_my_day(db_path: Path, employee: str, today: date, live_worklogs: list[dict] | None = None) -> dict:
    """Assemble the My Day payload; live_worklogs (issue_key, date, hours, summary) override this week's canonical hours."""
    run_id = resolve_canonical_run_id(db_path)
    if not run_id:
        raise RuntimeError("My Day needs a successful canonical refresh run.")
    who = _key(employee)
    week_start = _monday(today)
    history_start = week_start - timedelta(weeks=STREAK_WEEKS)
    horizon = today + timedelta(days=14)

    with closing(sqlite3.connect(db_path)) as conn, conn:
        profiles = _load_profiles(conn)
        state = conn.execute("SELECT updated_at_utc FROM canonical_refresh_state WHERE id = 1").fetchone()
        data_through = conn.execute("SELECT max(started_date) FROM canonical_worklogs WHERE run_id = ?", (run_id,)).fetchone()[0] or ""
        issue_rows = conn.execute(
            "SELECT issue_key, summary, status, issue_type, start_date, due_date, original_estimate_hours,"
            " total_hours_logged, epic_key FROM canonical_issues WHERE run_id = ? AND lower(trim(assignee)) = ?"
            " AND lower(issue_type) LIKE '%sub%task%' AND issue_key NOT LIKE 'RLT-%'", (run_id, who)).fetchall()
        log_rows = conn.execute(
            "SELECT w.issue_key, w.started_date, sum(w.hours_logged), coalesce(i.summary, ''), coalesce(i.epic_key, ''),"
            " coalesce(i.original_estimate_hours, 0), coalesce(i.assignee, '') FROM canonical_worklogs w"
            " LEFT JOIN canonical_issues i ON i.run_id = w.run_id AND i.issue_key = w.issue_key"
            " WHERE w.run_id = ? AND lower(trim(w.worklog_author)) = ? AND w.started_date BETWEEN ? AND ?"
            " AND w.issue_key NOT LIKE 'RLT-%' GROUP BY w.issue_key, w.started_date",
            (run_id, who, _iso(history_start), _iso(horizon))).fetchall()
        epic_keys = {row[8] for row in issue_rows if row[8]} | {row[4] for row in log_rows if row[4]}
        epic_names = dict(conn.execute(
            f"SELECT issue_key, summary FROM canonical_issues WHERE run_id = ? AND issue_key IN ({','.join('?' * len(epic_keys))})",
            (run_id, *epic_keys)).fetchall()) if epic_keys else {}

    logged: dict[tuple[str, str], float] = {}
    issue_meta: dict[str, dict] = {}
    for key, day, hours, summary, epic_key, estimate, assignee in log_rows:
        logged[(key, day)] = logged.get((key, day), 0.0) + float(hours or 0)
        issue_meta.setdefault(key, {"summary": summary, "epic_key": epic_key, "estimate": float(estimate or 0), "assignee": assignee})
    live = live_worklogs is not None
    if live:
        live_from = _iso(week_start)
        logged = {k: v for k, v in logged.items() if k[1] < live_from}
        for row in live_worklogs:
            k = (row["issue_key"], row["date"])
            logged[k] = logged.get(k, 0.0) + float(row["hours"])
            issue_meta.setdefault(row["issue_key"], {"summary": row.get("summary", ""), "epic_key": "", "estimate": 0.0, "assignee": ""})
        data_through = max(data_through, _iso(today))
    per_day: dict[str, float] = {}
    for (_, day), hours in logged.items():
        per_day[day] = per_day.get(day, 0.0) + hours

    leave_by_day: dict[str, float] = {}
    planned_leave: dict[str, dict] = {}
    for row in build_rlt_leave_snapshot(db_path, run_id, _iso(history_start), _iso(horizon)).get("daily") or []:
        if _key(row.get("assignee")) != who:
            continue
        day = str(row.get("period_day") or "")
        taken = sum(float(row.get(k) or 0) for k in ("planned_taken_hours", "unplanned_taken_hours", "unknown_taken_hours"))
        upcoming = float(row.get("planned_not_taken_hours") or 0)
        leave_by_day[day] = leave_by_day.get(day, 0.0) + taken + upcoming
        if upcoming > 0:
            planned_leave[day] = {"hours": upcoming, "jira": str(row.get("jira_task_ids") or "")}

    def describe(day: date) -> dict:
        iso = _iso(day)
        expected, kind = expected_hours(profiles, day)
        hours, leave = round(per_day.get(iso, 0.0), 2), round(leave_by_day.get(iso, 0.0), 2)
        status, missing = day_status(day, today, data_through, expected, hours, leave)
        return {"date": iso, "weekday": day.strftime("%a"), "expected": expected, "kind": kind,
                "logged": hours, "leave": leave, "status": status, "missing": missing}

    week = [describe(week_start + timedelta(days=i)) for i in range(5)]
    today_info = describe(today)

    missed = None
    probe = today - timedelta(days=1)
    for _ in range(10):
        info = describe(probe)
        if info["expected"] > 0 and info["status"] not in ("off",):
            if info["status"] == "missing":
                missed = info
            break
        probe -= timedelta(days=1)

    weeks: list[bool | None] = []
    for back in range(1, STREAK_WEEKS + 1):
        start = week_start - timedelta(weeks=back)
        days = [describe(start + timedelta(days=i)) for i in range(5)]
        if any(d["status"] == "unsynced" for d in days):
            weeks.append(None)
            continue
        weeks.append(all(d["status"] in ("complete", "leave", "off") for d in days))
    streak = 0
    for complete in weeks:
        if complete is None:
            continue
        if not complete:
            break
        streak += 1

    tasks = []
    today_iso = _iso(today)
    for key, summary, status, issue_type, start, due, estimate, total, epic_key in issue_rows:
        on_day = round(logged.get((key, today_iso), 0.0), 2)
        done = _key(status) in DONE_STATUSES
        scheduled = (not start or start <= today_iso) and (not due or due >= _iso(today - timedelta(days=14)))
        if (done or not scheduled) and not on_day:
            continue
        tasks.append({"issue_key": key, "summary": summary, "status": status, "issue_type": issue_type,
                      "start_date": start, "due_date": due, "estimate_hours": round(float(estimate or 0), 2),
                      "logged_hours": round(float(total or 0), 2), "logged_today": on_day, "done": done,
                      "overdue": bool(due and due < today_iso and not done),
                      "epic_key": epic_key, "epic_name": epic_names.get(epic_key, "")})
    tasks.sort(key=lambda t: (t["done"], -t["logged_today"], not t["overdue"], t["due_date"] or "9999", t["issue_key"]))

    assigned = {t["issue_key"] for t in tasks}
    recent_days: dict[str, str] = {}
    for (key, day) in logged:
        if day <= today_iso and day > recent_days.get(key, ""):
            recent_days[key] = day
    recent = [{"issue_key": key, "summary": issue_meta.get(key, {}).get("summary", ""), "last_logged": day,
               "epic_name": epic_names.get(issue_meta.get(key, {}).get("epic_key", ""), "")}
              for key, day in sorted(recent_days.items(), key=lambda kv: kv[1], reverse=True)
              if key not in assigned][:8]

    month_start = today.replace(day=1).isoformat()
    epics: dict[str, dict] = {}
    for (key, day), hours in logged.items():
        if not (month_start <= day <= today_iso):
            continue
        epic_key = issue_meta.get(key, {}).get("epic_key", "")
        if not epic_key:
            continue
        epic = epics.setdefault(epic_key, {"epic_key": epic_key, "epic_name": epic_names.get(epic_key, epic_key),
                                           "month_hours": 0.0, "planned_hours": 0.0, "logged_hours": 0.0})
        epic["month_hours"] += hours
    for key, _, _, _, _, _, estimate, total, epic_key in issue_rows:
        if epic_key in epics:
            epics[epic_key]["planned_hours"] += float(estimate or 0)
            epics[epic_key]["logged_hours"] += float(total or 0)
    for epic in epics.values():
        planned, spent = epic["planned_hours"], epic["logged_hours"]
        epic["state"] = "over" if planned and spent > planned + 0.01 else ("started" if not planned or spent / planned < 0.25 else "on_track")
        for field in ("month_hours", "planned_hours", "logged_hours"):
            epic[field] = round(epic[field], 2)
    top_epics = sorted(epics.values(), key=lambda e: e["month_hours"], reverse=True)[:3]

    upcoming = []
    for offset in range(1, 15):
        day = today + timedelta(days=offset)
        iso = _iso(day)
        if iso in planned_leave:
            upcoming.append({"date": iso, "kind": "leave", "title": "Planned leave",
                             "detail": f"{planned_leave[iso]['hours']:g}h · {planned_leave[iso]['jira']}".strip(" ·")})
        elif expected_hours(profiles, day)[1] == "holiday":
            upcoming.append({"date": iso, "kind": "holiday", "title": "Official holiday", "detail": "No hours expected"})
    upcoming = upcoming[:4]

    return {
        "employee": employee,
        "date": today_iso,
        "today": today_info,
        "missed": missed,
        "streak": {"weeks": streak, "recent": list(reversed(weeks[:6]))},
        "week": week,
        "week_total": round(sum(d["logged"] + d["leave"] for d in week), 2),
        "week_expected": round(sum(d["expected"] for d in week), 2),
        "tasks": tasks[:TASK_LIMIT],
        "recent": recent,
        "epics": top_epics,
        "upcoming": upcoming,
        "data_through": data_through,
        "refreshed_at": (state[0] if state else "") or "",
        "live": live,
        "canonical_run_id": run_id,
        "generated_at": datetime.now().astimezone().isoformat(),
    }
