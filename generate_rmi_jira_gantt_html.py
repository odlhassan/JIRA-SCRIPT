from __future__ import annotations

import argparse
import json
import os
from collections import defaultdict
from datetime import date, datetime, timezone
from html import escape
from pathlib import Path
from typing import Any

from canonical_report_data import build_rlt_leave_snapshot, load_canonical_issues, load_canonical_worklogs, resolve_canonical_run_id
from jira_client import BASE_URL
from generate_assignee_hours_report import DEFAULT_LEAVE_REPORT_INPUT_XLSX, _load_leave_daily_rows
from generate_employee_performance_report import _list_performance_teams, _load_performance_resource_resignation_map
from generate_nested_view_html import _load_capacity_profiles
from report_server import _load_epics_management_rows


DEFAULT_CAPACITY_DB = "assignee_hours_capacity.db"
DEFAULT_OUTPUT_HTML = "rmi_jira_gantt_report.html"
SECONDS_PER_HOUR = 3600
SECONDS_PER_DAY = 28800


def _resolve_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else base_dir / path


def _resolve_rmi_canonical_db_path(planner_db: Path) -> Path:
    """SQLite DB that holds canonical_issues / canonical_worklogs for this report.

    Defaults to the Epics Planner DB. Override with env ``JIRA_RMI_GANTT_CANONICAL_DB_PATH``
    (or CLI ``--canonical-db``) when canonical refresh writes to a different file.
    """
    planner_db = planner_db.resolve()
    raw = _to_text(os.environ.get("JIRA_RMI_GANTT_CANONICAL_DB_PATH", ""))
    if not raw:
        return planner_db
    alt = Path(raw)
    return alt.resolve() if alt.is_absolute() else (planner_db.parent / alt).resolve()


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _seconds_from_days(value: Any) -> float:
    return round(_to_float(value) * SECONDS_PER_DAY, 2)


def _seconds_from_hours(value: Any) -> float:
    return round(_to_float(value) * SECONDS_PER_HOUR, 2)


def _jira_url(issue_key: str, fallback: str = "") -> str:
    existing = _to_text(fallback)
    if existing:
        return existing
    key = _to_text(issue_key).upper()
    base = _to_text(BASE_URL).rstrip("/")
    return f"{base}/browse/{key}" if base and key else ""


def _issue_kind(issue_type: Any) -> str:
    text = _to_text(issue_type).lower()
    if "epic" in text:
        return "epic"
    if "story" in text:
        return "story"
    if "bug" in text:
        return "bug"
    if "sub-task" in text or "subtask" in text or "sub task" in text:
        return "subtask"
    return "issue"


def _format_generated_at() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def _parse_jira_datetime_for_display(started_date: str, started_utc: str) -> datetime | None:
    raw = _to_text(started_utc) or _to_text(started_date)
    if not raw:
        return None
    s = raw
    if len(s) >= 1 and s.endswith("Z") and "T" in s:
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        pass
    dpart = _to_text(started_date) or raw[:10]
    if len(dpart) >= 10 and dpart[4] == "-":
        try:
            return datetime.strptime(dpart[:10], "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return None


def _format_worklog_started_cell(started_date: str, started_utc: str) -> str:
    dt = _parse_jira_datetime_for_display(started_date, started_utc)
    if dt:
        return dt.strftime("%d-%b-%Y %H:%M")
    return _to_text(started_utc) or _to_text(started_date)


def _format_time_spent_jira_style(hours: Any) -> str:
    h = _to_float(hours)
    total_minutes = int(round(h * 60))
    if h > 0 and total_minutes == 0:
        total_minutes = 1
    hh, mm = divmod(max(0, total_minutes), 60)
    if hh and mm:
        return f"{hh}h {mm}m"
    if hh:
        return f"{hh}h"
    return f"{mm}m"


def _html_worklog_detail_rows(worklogs: list[dict[str, Any]]) -> str:
    parts: list[str] = []
    for w in worklogs:
        hrs = round(_to_float(w.get("hours_logged")), 2)
        parts.append(
            "<tr>"
            f"<td>{escape(_to_text(w.get('worklog_id')))}</td>"
            f"<td>{escape(_to_text(w.get('author')))}</td>"
            f"<td>{escape(_format_worklog_started_cell(_to_text(w.get('started_date')), _to_text(w.get('started_utc'))))}</td>"
            f"<td>{escape(_format_time_spent_jira_style(hrs))}</td>"
            f"<td>{escape(f'{hrs:.2f}')} h</td>"
            "</tr>"
        )
    return "".join(parts)


def _load_canonical_tree(db_path: Path, run_id: str) -> dict[str, Any]:
    issues = load_canonical_issues(db_path, run_id)
    worklogs = load_canonical_worklogs(db_path, run_id)

    by_key: dict[str, dict[str, Any]] = {}
    stories_by_epic: dict[str, list[dict[str, Any]]] = defaultdict(list)
    descendants_by_story: dict[str, list[dict[str, Any]]] = defaultdict(list)
    worklogs_by_issue: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for issue in issues:
        key = _to_text(issue.get("issue_key")).upper()
        if not key:
            continue
        by_key[key] = issue

    for worklog in worklogs:
        issue_key = _to_text(worklog.get("issue_key")).upper()
        if not issue_key:
            continue
        worklogs_by_issue[issue_key].append(
            {
                "worklog_id": _to_text(worklog.get("worklog_id")),
                "author": _to_text(worklog.get("worklog_author")),
                "started_date": _to_text(worklog.get("started_date")),
                "started_utc": _to_text(worklog.get("started_utc")),
                "hours_logged": round(_to_float(worklog.get("hours_logged")), 2),
            }
        )

    for issue in issues:
        issue_key = _to_text(issue.get("issue_key")).upper()
        if not issue_key:
            continue
        kind = _issue_kind(issue.get("issue_type"))
        epic_key = _to_text(issue.get("epic_key")).upper()
        parent_key = _to_text(issue.get("parent_issue_key")).upper()
        story_key = _to_text(issue.get("story_key")).upper()
        if kind == "story" and epic_key:
            stories_by_epic[epic_key].append(issue)
        elif kind in {"subtask", "bug", "issue"}:
            parent_story = story_key or parent_key
            if parent_story and parent_story != issue_key:
                descendants_by_story[parent_story].append(issue)

    for rows in stories_by_epic.values():
        rows.sort(key=lambda r: (_to_text(r.get("start_date")) or "9999", _to_text(r.get("due_date")) or "9999", _to_text(r.get("issue_key"))))
    for rows in descendants_by_story.values():
        rows.sort(key=lambda r: (_to_text(r.get("start_date")) or "9999", _to_text(r.get("due_date")) or "9999", _to_text(r.get("issue_key"))))

    return {
        "issues_by_key": by_key,
        "stories_by_epic": stories_by_epic,
        "descendants_by_story": descendants_by_story,
        "worklogs_by_issue": worklogs_by_issue,
    }


def _story_record(story: dict[str, Any], tree: dict[str, Any]) -> dict[str, Any]:
    story_key = _to_text(story.get("issue_key")).upper()
    descendants = []
    for child in tree["descendants_by_story"].get(story_key, []):
        issue_key = _to_text(child.get("issue_key")).upper()
        logs = tree["worklogs_by_issue"].get(issue_key, [])
        descendants.append(
            {
                "issue_key": issue_key,
                "title": _to_text(child.get("summary")),
                "issue_type": _to_text(child.get("issue_type")),
                "status": _to_text(child.get("status")),
                "priority": "",
                "assignee": _to_text(child.get("assignee")),
                "start_date": _to_text(child.get("start_date")),
                "due_date": _to_text(child.get("due_date")),
                "estimate_seconds": _seconds_from_hours(child.get("original_estimate_hours")),
                "logged_seconds": _seconds_from_hours(sum(_to_float(w.get("hours_logged")) for w in logs)),
                "jira_url": _jira_url(issue_key),
                "worklogs": logs,
            }
        )
    logs = tree["worklogs_by_issue"].get(story_key, [])
    return {
        "story_key": story_key,
        "title": _to_text(story.get("summary")),
        "issue_type": _to_text(story.get("issue_type")),
        "status": _to_text(story.get("status")),
        "priority": "",
        "assignee": _to_text(story.get("assignee")),
        "start_date": _to_text(story.get("start_date")),
        "due_date": _to_text(story.get("due_date")),
        "estimate_seconds": _seconds_from_hours(story.get("original_estimate_hours")),
        "logged_seconds": _seconds_from_hours(sum(_to_float(w.get("hours_logged")) for w in logs)),
        "jira_url": _jira_url(story_key),
        "worklogs": logs,
        "subtasks": descendants,
    }


def _epic_metrics(epic: dict[str, Any]) -> dict[str, float]:
    plan = epic.get("plans", {}).get("epic_plan", {}) if isinstance(epic.get("plans"), dict) else {}
    return {
        "most_likely_seconds": _seconds_from_days(plan.get("most_likely_man_days") if plan.get("most_likely_man_days") not in (None, "") else plan.get("man_days")),
        "optimistic_seconds": _seconds_from_days(plan.get("optimistic_man_days")),
        "pessimistic_seconds": _seconds_from_days(plan.get("pessimistic_man_days")),
        "calculated_seconds": _seconds_from_days(plan.get("calculated_man_days")),
        "tk_approved_seconds": _seconds_from_days(plan.get("tk_approved_man_days") if plan.get("tk_approved_man_days") not in (None, "") else plan.get("tk_budgeted_man_days")),
    }


def load_report_data(db_path: Path, run_id: str = "") -> dict[str, Any]:
    planner_db = Path(db_path).resolve()
    canonical_db = _resolve_rmi_canonical_db_path(planner_db)
    requested_run = _to_text(run_id)
    effective_run_id = resolve_canonical_run_id(canonical_db, requested_run)
    if not effective_run_id and canonical_db != planner_db:
        effective_run_id = resolve_canonical_run_id(planner_db, requested_run)
    tree = _load_canonical_tree(canonical_db, effective_run_id) if effective_run_id else {
        "issues_by_key": {},
        "stories_by_epic": {},
        "descendants_by_story": {},
        "worklogs_by_issue": {},
    }
    planner_rows = [row for row in _load_epics_management_rows(planner_db) if int(row.get("is_tk_epic") or 0) == 1]
    epic_records: list[dict[str, Any]] = []

    for row in planner_rows:
        epic_key = _to_text(row.get("epic_key")).upper()
        if not epic_key:
            continue
        product = _to_text(row.get("project_name")) or _to_text(row.get("project_key")) or "Unassigned"
        canonical_epic = tree["issues_by_key"].get(epic_key, {})
        epic_plan = row.get("plans", {}).get("epic_plan", {}) if isinstance(row.get("plans"), dict) else {}
        start_date = _to_text(epic_plan.get("start_date")) or _to_text(canonical_epic.get("start_date"))
        due_date = _to_text(epic_plan.get("due_date")) or _to_text(canonical_epic.get("due_date"))
        stories = [_story_record(story, tree) for story in tree["stories_by_epic"].get(epic_key, [])]
        story_estimate_seconds = sum(_to_float(story.get("estimate_seconds")) for story in stories)
        subtask_estimate_seconds = sum(
            _to_float(child.get("estimate_seconds"))
            for story in stories
            for child in story.get("subtasks", [])
        )
        logged_seconds = sum(
            _to_float(story.get("logged_seconds")) + sum(_to_float(child.get("logged_seconds")) for child in story.get("subtasks", []))
            for story in stories
        )
        metrics = _epic_metrics(row)
        epic_records.append(
            {
                "jira_id": epic_key,
                "title": _to_text(row.get("epic_name")) or _to_text(canonical_epic.get("summary")) or epic_key,
                "product": product,
                "project_key": _to_text(row.get("project_key")),
                "component": _to_text(row.get("component")),
                "status": _to_text(canonical_epic.get("status")) or _to_text(row.get("delivery_status")),
                "priority": _to_text(row.get("priority")),
                "start_date": start_date,
                "due_date": due_date,
                "jira_url": _jira_url(epic_key, _to_text(row.get("jira_url"))),
                "jira_populated": bool(canonical_epic or stories),
                "story_count": len(stories),
                "story_estimate_seconds": round(story_estimate_seconds, 2),
                "subtask_estimate_seconds": round(subtask_estimate_seconds, 2),
                "logged_seconds": round(logged_seconds, 2),
                "jira_original_estimate_seconds": _seconds_from_hours(canonical_epic.get("original_estimate_hours")),
                "epics_planner_remarks": _to_text(row.get("remarks")),
                "stories": stories,
                **metrics,
            }
        )

    epic_records.sort(key=lambda r: (str(r.get("product", "")).lower(), r.get("start_date") or "9999", str(r.get("jira_id", "")).lower()))
    rmi_schedule_records = build_rmi_schedule_records(epic_records)
    return {
        "generated_at": _format_generated_at(),
        "database_path": str(planner_db),
        "canonical_database_path": str(canonical_db),
        "canonical_run_id": effective_run_id,
        "epics": epic_records,
        "summary": build_summary(epic_records),
        "metric_summary": build_metric_summary(epic_records),
        "capacity_source": build_capacity_source(epic_records, planner_db, effective_run_id),
        "rmi_schedule_records": rmi_schedule_records,
        "rmi_schedule_years": build_rmi_schedule_years(rmi_schedule_records),
    }


def build_rmi_schedule_records(epics: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for epic in epics:
        stories_out: list[dict[str, Any]] = []
        for story in epic.get("stories") or []:
            subs_out = []
            for sub in story.get("subtasks") or []:
                subs_out.append(
                    {
                        "start_date": _to_text(sub.get("start_date")),
                        "due_date": _to_text(sub.get("due_date")),
                        "estimate_seconds": round(_to_float(sub.get("estimate_seconds")), 4),
                    }
                )
            stories_out.append(
                {
                    "start_date": _to_text(story.get("start_date")),
                    "due_date": _to_text(story.get("due_date")),
                    "estimate_seconds": round(_to_float(story.get("estimate_seconds")), 4),
                    "subtasks": subs_out,
                }
            )
        records.append(
            {
                "roadmap_item": _to_text(epic.get("title")),
                "jira_id": _to_text(epic.get("jira_id")).upper(),
                "jira_url": _to_text(epic.get("jira_url")),
                "jira_populated": bool(epic.get("jira_populated")),
                "product": _to_text(epic.get("product")) or "Unassigned",
                "status": _to_text(epic.get("status")),
                "start_date": _to_text(epic.get("start_date")),
                "due_date": _to_text(epic.get("due_date")),
                "most_likely_days": round(_to_float(epic.get("most_likely_seconds")) / SECONDS_PER_DAY, 8),
                "tk_approved_days": round(_to_float(epic.get("tk_approved_seconds")) / SECONDS_PER_DAY, 8),
                "stories": stories_out,
            }
        )
    return records


def build_rmi_schedule_years(records: list[dict[str, Any]]) -> list[int]:
    years: set[int] = {datetime.now(timezone.utc).year}

    def ingest(iso: Any) -> None:
        t = _to_text(iso)
        if len(t) >= 4 and t[0:4].isdigit():
            years.add(int(t[0:4]))

    for r in records:
        ingest(r.get("start_date"))
        ingest(r.get("due_date"))
        for s in r.get("stories") or []:
            ingest(s.get("start_date"))
            ingest(s.get("due_date"))
            for c in s.get("subtasks") or []:
                ingest(c.get("start_date"))
                ingest(c.get("due_date"))
    return sorted(years)


def build_summary(epics: list[dict[str, Any]]) -> dict[str, Any]:
    story_count = sum(len(epic.get("stories", [])) for epic in epics)
    subtask_count = sum(len(story.get("subtasks", [])) for epic in epics for story in epic.get("stories", []))
    worklog_count = sum(
        len(story.get("worklogs", [])) + sum(len(child.get("worklogs", [])) for child in story.get("subtasks", []))
        for epic in epics
        for story in epic.get("stories", [])
    )
    return {
        "epic_count": len(epics),
        "story_count": story_count,
        "descendant_count": subtask_count,
        "worklog_count": worklog_count,
        "total_worklog_seconds": round(sum(_to_float(epic.get("logged_seconds")) for epic in epics), 2),
    }


def build_metric_summary(epics: list[dict[str, Any]]) -> dict[str, dict[str, float]]:
    metric_keys = [
        "epic_count",
        "most_likely_seconds",
        "optimistic_seconds",
        "pessimistic_seconds",
        "calculated_seconds",
        "tk_approved_seconds",
        "jira_original_estimate_seconds",
        "story_estimate_seconds",
        "subtask_estimate_seconds",
        "logged_seconds",
    ]

    def empty() -> dict[str, float]:
        return {key: 0.0 for key in metric_keys}

    out: dict[str, dict[str, float]] = {"all": empty()}
    for epic in epics:
        product = _to_text(epic.get("product")) or "Unassigned"
        totals_list = [out["all"], out.setdefault(product, empty())]
        for totals in totals_list:
            totals["epic_count"] += 1
            for key in metric_keys:
                if key == "epic_count":
                    continue
                totals[key] += _to_float(epic.get(key))
    return out


def build_capacity_source(epics: list[dict[str, Any]], db_path: Path, run_id: str = "") -> dict[str, Any]:
    employees_by_product: dict[str, set[str]] = defaultdict(set)

    def add_assignee(product_name: str, assignee_value: Any) -> None:
        assignee = _to_text(assignee_value)
        if not assignee or assignee.lower() == "unassigned":
            return
        product = _to_text(product_name) or "Unassigned"
        employees_by_product["all"].add(assignee)
        employees_by_product[product].add(assignee)

    for epic in epics:
        product = _to_text(epic.get("product")) or "Unassigned"
        for story in epic.get("stories", []):
            add_assignee(product, story.get("assignee"))
            for child in story.get("subtasks", []):
                add_assignee(product, child.get("assignee"))

    # Planned leave hours: match nested view scorecard (rlt_leave_report.xlsx Daily_Assignee):
    # totalPlannedLeaves = planned_taken_hours + planned_not_taken_hours per day row.
    leaves_hours_by_month: dict[str, float] = defaultdict(float)
    leaves_hours_by_month_assignee: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    base_dir = db_path.parent
    leave_name = os.getenv("JIRA_LEAVE_REPORT_XLSX_PATH", DEFAULT_LEAVE_REPORT_INPUT_XLSX).strip() or DEFAULT_LEAVE_REPORT_INPUT_XLSX
    leave_path = _resolve_path(leave_name, base_dir)
    for row in _load_leave_daily_rows(leave_path):
        period_day = _to_text(row.get("period_day"))
        if len(period_day) < 7 or period_day[4] != "-":
            continue
        month = period_day[:7]
        taken = _to_float(row.get("planned_taken_hours"))
        not_taken = _to_float(row.get("planned_not_taken_hours"))
        planned = taken + not_taken
        if planned <= 0:
            continue
        assignee = _to_text(row.get("assignee"))
        leaves_hours_by_month[month] += planned
        if assignee:
            leaves_hours_by_month_assignee[month][assignee] += planned

    leave_snapshot = build_rlt_leave_snapshot(db_path, run_id)
    rlt_employees: list[str] = sorted(
        {_to_text(row.get("assignee")) for row in leave_snapshot.get("assignee_summary", []) if _to_text(row.get("assignee"))},
        key=lambda v: v.lower(),
    )

    # Load performance teams for team selector + employee breakdown drawer
    teams: list[dict[str, Any]] = []
    assignee_to_teams: dict[str, list[str]] = defaultdict(list)
    member_records: dict[str, dict[str, Any]] = {}
    try:
        raw_teams = _list_performance_teams(db_path)
        for team in raw_teams:
            tn = _to_text(team.get("team_name"))
            assignee_list: list[str] = []
            for a in team.get("assignees") or []:
                a = _to_text(a)
                if not a:
                    continue
                assignee_list.append(a)
                if tn and tn not in assignee_to_teams[a]:
                    assignee_to_teams[a].append(tn)
            teams.append(
                {
                    "team_name": tn,
                    "team_leader": _to_text(team.get("team_leader")),
                    "assignees": assignee_list,
                }
            )
        for a in assignee_to_teams:
            assignee_to_teams[a].sort(key=lambda x: str(x).lower())
        if assignee_to_teams:
            member_records = _load_performance_resource_resignation_map(
                db_path,
                sorted(assignee_to_teams.keys(), key=lambda v: str(v).lower()),
            )
    except Exception:
        pass

    org_team_assignees: list[str] = sorted(set(assignee_to_teams.keys()), key=lambda v: v.lower()) if assignee_to_teams else []
    rmi_all_assignees: list[str] = sorted(
        employees_by_product.get("all", set()),
        key=lambda v: str(v).lower(),
    )
    def _dedupe_names(*iterables: list[str]) -> list[str]:
        seen: set[str] = set()
        out: list[str] = []
        for it in iterables:
            for x in it:
                t = _to_text(x)
                if not t or t.lower() == "unassigned":
                    continue
                k = t.lower()
                if k in seen:
                    continue
                seen.add(k)
                out.append(t)
        out.sort(key=str.lower)
        return out

    capacity_universe_assignees: list[str] = _dedupe_names(org_team_assignees, rlt_employees, rmi_all_assignees)

    capacity_profiles = _load_capacity_profiles(db_path)

    return {
        "employees_by_product": {
            product: sorted(assignees, key=lambda value: value.lower())
            for product, assignees in employees_by_product.items()
        },
        "org_team_assignees": org_team_assignees,
        "assignee_to_teams": {k: v for k, v in assignee_to_teams.items() if v},
        "capacity_universe_assignees": capacity_universe_assignees,
        "rlt_employees": rlt_employees,
        "teams": teams,
        "member_records": member_records,
        "capacity_profiles": capacity_profiles,
        "leaves_hours_by_month": {
            month: round(hours, 2)
            for month, hours in sorted(leaves_hours_by_month.items())
        },
        "leaves_hours_by_month_assignee": {
            month: {
                assignee: round(hours, 2)
                for assignee, hours in sorted(by_assignee.items(), key=lambda item: item[0].lower())
            }
            for month, by_assignee in sorted(leaves_hours_by_month_assignee.items())
        },
    }



def _safe_json_for_script(payload_json: str) -> str:
    # Inside <script>, HTML entities are NOT decoded. We must NOT html-escape JSON.
    # We only need to neutralise sequences that could prematurely close the script
    # tag or be reinterpreted by the HTML parser.
    return (
        payload_json
        .replace("<!--", "<\\!--")
        .replace("</", "<\\/")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def _json_script(id_value: str, value: Any) -> str:
    payload = _safe_json_for_script(json.dumps(value, ensure_ascii=True))
    return f'<script type="application/json" id="{escape(id_value)}">{payload}</script>'


def _duration(seconds: Any, unit: str = "hours") -> str:
    value = _to_float(seconds)
    if unit == "days":
        return f"{value / SECONDS_PER_DAY:,.2f} d"
    return f"{value / SECONDS_PER_HOUR:,.2f} h"


def _available_months(epics: list[dict[str, Any]]) -> list[str]:
    months = set()
    for epic in epics:
        for field in ("start_date", "due_date"):
            text = _to_text(epic.get(field))
            if len(text) >= 7 and text[4] == "-":
                months.add(text[:7])
    return sorted(months)



# ── Rich Report CSS ──────────────────────────────────────────────────

_REPORT_CSS = """
:root {
  --bg: #e8eef5;
  --panel: #ffffff;
  --panel-soft: #f7fafd;
  --text: #0d1b2a;
  --muted: #516174;
  --line: #d0dbe6;
  --shadow: 0 1px 2px rgba(16,32,51,0.04), 0 4px 12px rgba(16,32,51,0.06), 0 16px 40px rgba(16,32,51,0.07);
  --shadow-sm: 0 1px 2px rgba(16,32,51,0.04), 0 4px 10px rgba(16,32,51,0.05);
  --shadow-lg: 0 2px 4px rgba(16,32,51,0.04), 0 8px 24px rgba(16,32,51,0.08), 0 28px 56px rgba(16,32,51,0.10);
  --radius-sm: 10px;
  --radius-md: 16px;
  --radius-lg: 22px;
  --ring: rgba(37,99,235,0.26);
  --gutter: clamp(16px,2.5vw,40px);
  --row-epic: #f0e8ff;
  --row-epic-hover: #e6d8ff;
  --row-story: #ddf0fd;
  --row-story-hover: #c9e7fb;
  --row-subtask: #d8fae8;
  --row-subtask-hover: #c2f5d6;
  --row-bug: #fde8e8;
  --row-bug-hover: #fcd0d0;
}
*, *::before, *::after { box-sizing: border-box; }
body {
  margin: 0;
  font-family: "Inter","Segoe UI",system-ui,-apple-system,sans-serif;
  -webkit-font-smoothing: antialiased;
  color: var(--text);
  background:
    radial-gradient(ellipse 90% 50% at top left,rgba(15,118,110,.09),transparent 42%),
    radial-gradient(ellipse 70% 40% at top right,rgba(190,24,93,.08),transparent 40%),
    linear-gradient(180deg,#f2f7fc 0%,var(--bg) 100%);
}
.page { width:100%; max-width:none; margin:0; padding:28px var(--gutter) 44px; }
header { margin-bottom:18px; }
h1 { margin:0 0 8px; font-size:2.2rem; letter-spacing:-0.04em; font-weight:800; }
h2 { margin:0 0 10px; font-size:1.2rem; font-weight:700; }
.subtext { color:var(--muted); line-height:1.55; max-width:1160px; font-size:.94rem; }

/* ── Metric cards ─────────────────────────────────────────────────── */
.metric-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; margin:18px 0 22px; }
.metric-card {
  padding:18px 20px; position:relative; overflow:hidden;
  background:var(--panel); border:1px solid rgba(255,255,255,.16);
  border-radius:var(--radius-lg); box-shadow:var(--shadow); color:#fff;
}
.metric-card::before {
  content:""; position:absolute; inset:auto -18px -38px auto;
  width:112px; height:112px; border-radius:999px;
  background:rgba(255,255,255,.14); pointer-events:none;
}
.metric-card-teal    { background:linear-gradient(135deg,#0f766e,#14b8a6); }
.metric-card-blue    { background:linear-gradient(135deg,#1d4ed8,#38bdf8); }
.metric-card-amber   { background:linear-gradient(135deg,#b45309,#f59e0b); }
.metric-card-rose    { background:linear-gradient(135deg,#be123c,#fb7185); }
.metric-card-violet  { background:linear-gradient(135deg,#6d28d9,#a78bfa); }
.metric-card-cyan    { background:linear-gradient(135deg,#0f766e,#67e8f9); color:#062a30; }
.metric-card-emerald { background:linear-gradient(135deg,#047857,#34d399); }
.metric-card-indigo  { background:linear-gradient(135deg,#4338ca,#818cf8); }
.metric-card-slate   { background:linear-gradient(135deg,#334155,#64748b); }
.metric-card-capacity { background:linear-gradient(135deg,#0f766e,#14b8a6); }
.metric-label { color:rgba(255,255,255,.82); font-size:.92rem; margin-bottom:8px; }
.metric-value-wrap { min-height:42px; display:flex; align-items:center; }
.metric-value { font-size:1.9rem; font-weight:700; letter-spacing:-0.03em; }
.metric-meta { color:rgba(255,255,255,.82); margin-top:6px; font-size:.92rem; max-width:28ch; }
.metric-card-clickable { cursor:pointer; }
.metric-card-clickable:hover { filter:brightness(1.06); }
.metric-click-icon {
  display:inline-flex; align-items:center; justify-content:center;
  float:right; width:28px; height:28px; border-radius:50%;
  background:rgba(255,255,255,.22); font-size:.82rem; margin-top:-2px;
}

/* ── Estimate range group ─────────────────────────────────────────── */
.estimate-cards-group {
  grid-column:1/-1; display:grid; grid-template-columns:repeat(4,minmax(0,1fr));
  gap:0; border-radius:18px; overflow:hidden; box-shadow:var(--shadow);
  border:1px solid rgba(29,78,216,.18); position:relative;
}
.estimate-cards-group::before {
  content:"Estimation Range"; position:absolute; top:10px; left:14px;
  font-size:.64rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase;
  color:rgba(255,255,255,.68); z-index:2; pointer-events:none;
}
.estimate-cards-group .metric-card { border-radius:0; box-shadow:none; border:0; padding:28px 20px 20px; }
.estimate-cards-group .metric-card::before { display:none; }
.estimate-cards-group .metric-card+.metric-card { border-left:1px solid rgba(255,255,255,.18); }
.estimate-cards-group .metric-card::after {
  content:""; position:absolute; top:44px; right:-9px; width:18px; height:18px;
  border-top:2px solid rgba(255,255,255,.45); border-right:2px solid rgba(255,255,255,.45);
  transform:rotate(45deg); pointer-events:none; z-index:3;
}
.estimate-cards-group .metric-card:last-child::after { display:none; }
.metric-estimate-step-1 { background:linear-gradient(180deg,#7cb2fb,#60a5fa); }
.metric-estimate-step-2 { background:linear-gradient(180deg,#4f93f7,#3b82f6); }
.metric-estimate-step-3 { background:linear-gradient(180deg,#2f6fe6,#1d4ed8); }
.metric-estimate-step-4 { background:linear-gradient(180deg,#1d3fae,#1e3a8a); }

/* ── Hero card ────────────────────────────────────────────────────── */
.metric-card-hero { grid-column:span 2; padding:22px 26px 24px; min-height:156px; }
.metric-card-hero .metric-label { font-size:1rem; font-weight:800; margin-bottom:10px; }
.metric-card-hero .metric-value { font-size:2.6rem; font-weight:800; }
.metric-card-hero::before { width:150px; height:150px; inset:auto -24px -50px auto; }

/* ── Product summary cards ────────────────────────────────────────── */
.product-summary-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; margin:-4px 0 22px; }
.product-summary-card {
  background:linear-gradient(180deg,#fff,#f7fafd);
  border:1px solid rgba(200,214,228,.80); border-left:5px solid var(--product-accent);
  border-radius:var(--radius-lg); box-shadow:var(--shadow);
  padding:16px 18px; cursor:pointer; transition:transform .15s,box-shadow .15s;
}
.product-summary-card:hover { transform:translateY(-2px); box-shadow:var(--shadow-lg); }
.product-summary-card.active {
  border-color:color-mix(in srgb,var(--product-accent) 45%,white);
  background:linear-gradient(180deg,color-mix(in srgb,var(--product-accent) 10%,white),color-mix(in srgb,var(--product-accent) 4%,#f8fbff));
}
.product-summary-label { color:var(--text); font-size:.92rem; font-weight:800; margin-bottom:8px; }
.product-summary-duration { font-size:1.55rem; font-weight:800; color:var(--product-accent); letter-spacing:-0.03em; }
.product-summary-meta { color:var(--muted); font-size:.88rem; margin-top:4px; }

/* ── Panels ───────────────────────────────────────────────────────── */
.panel {
  background:var(--panel); border:1px solid rgba(200,214,228,.80);
  border-radius:var(--radius-lg); box-shadow:var(--shadow);
  padding:20px 22px 24px; margin-bottom:20px;
}

/* ── Capacity calculator ──────────────────────────────────────────── */
.capacity-calculator {
  display:grid; grid-template-columns:auto 1fr; gap:0;
  border:1px solid rgba(200,214,228,.80); border-radius:var(--radius-lg);
  background:var(--panel); box-shadow:var(--shadow); margin-bottom:20px;
  overflow:visible;
}
/* Lift above following sections while team dropdown is open (panel extends below this card). */
.capacity-calculator.capacity-calculator--dropdown-open {
  position:relative; z-index:80;
}
.capacity-calc-title {
  writing-mode:vertical-rl; transform:rotate(180deg);
  padding:16px 10px; font-size:.82rem; font-weight:800; text-transform:uppercase;
  letter-spacing:.14em; color:#fff; background:linear-gradient(180deg,#0f766e,#14b8a6);
  display:flex; align-items:center; justify-content:center;
  border-radius:var(--radius-lg) 0 0 var(--radius-lg);
}
.capacity-calc-body {
  display:flex; flex-wrap:wrap; gap:14px; padding:18px 20px; align-items:flex-end;
  border-radius:0 var(--radius-lg) var(--radius-lg) 0;
}
.capacity-field { display:grid; gap:5px; }
.capacity-field-label { font-size:.82rem; font-weight:800; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; }
.capacity-field-value {
  padding:10px 14px; border:1px solid rgba(192,206,218,.90); border-radius:var(--radius-sm);
  background:#f8fbff; color:var(--text); font:inherit; font-weight:700;
  box-shadow:0 1px 3px rgba(16,32,51,.06); min-width:100px;
}
.capacity-results-grid {
  width:100%;
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:14px;
}
.capacity-result-card { width:auto; min-width:0; padding:14px 16px; height:100%; }
/* ── Capacity team multi-select ─────────────────────────────────────── */
.capacity-team-field { position:relative; z-index:1; min-width:200px; max-width:320px; }
.capacity-ms { position:relative; }
.capacity-ms.capacity-ms-open { z-index:60; }
.capacity-ms-trigger {
  appearance:none; display:flex; align-items:center; justify-content:space-between; gap:10px;
  width:100%; padding:9px 14px; border:1px solid rgba(192,206,218,.90);
  border-radius:var(--radius-sm); background:#fff;
  color:var(--text); font:inherit; font-weight:600; cursor:pointer;
  box-shadow:0 1px 3px rgba(16,32,51,.06); text-align:left;
}
.capacity-ms-trigger:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 3px var(--ring); }
.capacity-ms-trigger-label { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.capacity-ms-chevron {
  flex-shrink:0; width:12px; height:8px; opacity:.75;
  background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23516174' stroke-width='1.5' fill='none'/%3E%3C/svg%3E") no-repeat center;
  transition:transform .18s ease;
}
.capacity-ms-open .capacity-ms-chevron { transform:rotate(180deg); }
.capacity-ms-panel {
  position:absolute; left:0; top:calc(100% + 4px); min-width:100%; width:max-content; max-width:min(380px,92vw);
  background:#fff; border:1px solid rgba(192,206,218,.95); border-radius:var(--radius-sm);
  box-shadow:0 10px 28px rgba(16,32,51,.14); overflow:hidden;
}
.capacity-ms-search-wrap { padding:8px 10px; border-bottom:1px solid rgba(200,214,228,.85); background:#f8fbff; }
.capacity-ms-search {
  width:100%; box-sizing:border-box; padding:8px 10px; border:1px solid rgba(192,206,218,.90);
  border-radius:var(--radius-sm); font:inherit; color:var(--text); background:#fff;
}
.capacity-ms-search:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 2px var(--ring); }
.capacity-ms-list { max-height:240px; overflow-y:auto; padding:4px 0; margin:0; }
.capacity-ms-item {
  display:flex; align-items:flex-start; gap:10px; margin:0; padding:8px 12px; cursor:pointer;
  font:inherit; font-weight:600; color:var(--text); user-select:none;
}
.capacity-ms-item:hover, .capacity-ms-item:focus-within { background:rgba(37,99,235,.06); }
.capacity-ms-item input { margin-top:3px; accent-color:#0f766e; cursor:pointer; flex-shrink:0; }
.capacity-ms-item-text { flex:1; line-height:1.35; }
.capacity-ms-item-all { border-bottom:1px solid rgba(200,214,228,.65); margin-bottom:2px; padding-bottom:10px; }
.capacity-ms-team-group { border-bottom:1px solid rgba(200,214,228,.45); }
.capacity-ms-team-group:last-child { border-bottom:none; }
.capacity-ms-team-row { font-weight:700; }
.capacity-ms-members { padding:0 0 6px 0; }
.capacity-ms-member-row { padding-left:32px; font-weight:500; }
.capacity-ms-hidden { display:none !important; }
.capacity-ms-actions { padding:6px 10px 8px; border-top:1px solid rgba(200,214,228,.75); background:#fafcff; }
.capacity-ms-link {
  appearance:none; border:none; background:none; padding:0; margin:0; cursor:pointer;
  font:inherit; font-weight:700; font-size:.84rem; color:#2563eb;
}
.capacity-ms-link:hover { text-decoration:underline; }

/* ── Search toolbar ───────────────────────────────────────────────── */
.search-toolbar {
  display:flex; gap:12px; flex-wrap:wrap; align-items:center;
  margin:18px 0 12px; padding:14px 18px;
  border:1px solid rgba(200,214,228,.75); border-radius:var(--radius-md);
  background:rgba(255,255,255,.90); box-shadow:var(--shadow-sm);
}
.search-toolbar-label { font-size:.82rem; font-weight:800; text-transform:uppercase; letter-spacing:.04em; color:#52657a; }
.search-input {
  flex:1 1 360px; min-width:240px; padding:11px 14px;
  border:1px solid rgba(192,206,218,.90); border-radius:var(--radius-sm);
  background:#fff; color:var(--text); font:inherit;
}
.search-input:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 3px var(--ring); }
.search-clear {
  appearance:none; border:1px solid rgba(192,206,218,.90); background:#fff; color:#203141;
  border-radius:999px; padding:9px 14px; font-size:.84rem; font-weight:700; cursor:pointer;
}
.search-status { min-height:1.25rem; margin:0 0 12px; color:#5c6f83; font-size:.9rem; font-weight:700; }

/* ── View/unit toggle ─────────────────────────────────────────────── */
.view-toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:10px 0 18px; }
.view-toggle, .unit-toggle {
  appearance:none; border:1px solid rgba(200,214,228,.90); background:#fff; color:var(--text);
  border-radius:999px; padding:10px 16px; font-weight:700; cursor:pointer;
  box-shadow:0 1px 3px rgba(16,32,51,.07); transition:all .15s;
}
.view-toggle:hover, .unit-toggle:hover { transform:translateY(-1px); box-shadow:0 4px 12px rgba(16,32,51,.10); }
.view-toggle.active { background:#0d1b2a; color:#fff; border-color:#0d1b2a; }
.unit-toggle.active { background:#1d4ed8; color:#fff; border-color:#1d4ed8; }

/* ── Product toolbar ──────────────────────────────────────────────── */
.product-toolbar {
  display:flex; gap:8px; flex-wrap:wrap; margin:0 0 12px; padding:10px 14px;
  border:1px solid rgba(200,214,228,.80); border-radius:var(--radius-sm);
  background:rgba(248,250,255,.92);
}
.product-toolbar-label {
  align-self:center; margin-right:6px; font-size:.82rem; font-weight:800;
  text-transform:uppercase; letter-spacing:.04em; color:#52657a;
}
.product-toggle {
  appearance:none; border:1px solid rgba(192,206,218,.90); background:#fff; color:#203141;
  border-radius:999px; padding:7px 12px; font-size:.84rem; font-weight:700; cursor:pointer;
  transition:all .12s;
}
.product-toggle.active { background:#0d1b2a; color:#fff; border-color:#0d1b2a; }

/* ── TK Month filter ──────────────────────────────────────────────── */
.tk-month-toolbar {
  display:flex; align-items:center; gap:14px; flex-wrap:wrap;
  margin:-4px 0 20px; padding:14px 18px;
  border:1px solid rgba(200,214,228,.75); border-radius:var(--radius-md);
  background:rgba(255,255,255,.90); box-shadow:var(--shadow-sm);
}
.tk-month-toggle { display:inline-flex; align-items:center; gap:10px; cursor:pointer; user-select:none; font-weight:700; color:#2a3a4d; font-size:.92rem; }
.tk-month-toggle-input { position:absolute; opacity:0; width:0; height:0; }
.tk-month-toggle-track { position:relative; display:inline-block; width:42px; height:24px; background:#c2ceda; border-radius:999px; transition:background .18s; flex-shrink:0; }
.tk-month-toggle-thumb { position:absolute; top:3px; left:3px; width:18px; height:18px; border-radius:50%; background:#fff; box-shadow:0 1px 3px rgba(15,23,42,.20); transition:transform .22s cubic-bezier(.34,1.56,.64,1); }
.tk-month-toggle-input:checked+.tk-month-toggle-track { background:#2563eb; }
.tk-month-toggle-input:checked+.tk-month-toggle-track .tk-month-toggle-thumb { transform:translateX(18px); }
.tk-month-select {
  appearance:none; padding:9px 32px 9px 14px; border:1px solid rgba(192,206,218,.90);
  border-radius:var(--radius-sm); background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23516174' stroke-width='1.5' fill='none'/%3E%3C/svg%3E") no-repeat right 12px center;
  color:var(--text); font:inherit; font-weight:600; min-width:140px; cursor:pointer;
}
.tk-month-select:disabled { opacity:.55; cursor:not-allowed; }
.tk-month-status { color:#52657a; font-size:.88rem; font-weight:600; }

/* ── Month analysis panel ─────────────────────────────────────────── */
.tk-month-analysis-panel {
  border-color:rgba(29,78,216,.18);
  background:radial-gradient(circle at top left,rgba(14,165,233,.10),transparent 24%),linear-gradient(180deg,#fff,#f8fbff);
}
.tk-month-analysis-header { display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:12px; margin-bottom:16px; }
.tk-month-chart-bars { display:grid; grid-template-columns:repeat(3,1fr); gap:16px; align-items:end; min-height:280px; }
.tk-month-chart-bars.tk-month-chart-bars--all {
  display:flex; flex-direction:row; flex-wrap:nowrap; align-items:flex-end; gap:12px; min-height:280px;
  overflow-x:auto; -webkit-overflow-scrolling:touch; padding-bottom:4px;
}
.tk-month-chart-bars.tk-month-chart-bars--all .tk-month-chart-bar-card {
  flex:0 0 auto; min-width:100px; max-width:148px; min-height:240px;
}
.tk-month-chart-bar-card { min-height:240px; display:flex; flex-direction:column; align-items:center; gap:10px; padding-top:10px; }
.tk-month-chart-value { color:#0f172a; font-size:1rem; font-weight:800; }
.tk-month-chart-bar-track {
  width:min(100%,148px); height:180px; padding:10px; display:flex; align-items:flex-end;
  border:1px solid rgba(203,213,225,.95); border-radius:18px 18px 10px 10px;
  background:linear-gradient(180deg,rgba(241,245,249,.85),rgba(255,255,255,.98));
}
.tk-month-chart-bar-fill {
  width:100%; min-height:6px; border-radius:12px 12px 8px 8px;
  background:linear-gradient(180deg,#60a5fa,#2563eb); box-shadow:0 12px 22px rgba(37,99,235,.22);
  transition:height .25s;
}
.tk-month-chart-bar-card.featured .tk-month-chart-bar-fill {
  background:linear-gradient(180deg,#34d399,#059669); box-shadow:0 12px 22px rgba(5,150,105,.24);
}
.tk-month-chart-label { color:#334155; font-size:.9rem; font-weight:800; text-align:center; }
.tk-month-analysis-summary { display:flex; flex-wrap:wrap; gap:14px; margin:16px 0 18px; align-items:stretch; }
.tk-month-analysis-card {
  flex:1 1 320px; min-width:280px; max-width:100%; padding:12px 14px 14px; border-radius:14px;
  border:1px solid rgba(191,219,254,.95); background:rgba(239,246,255,.92);
  display:flex; flex-direction:column; gap:8px;
}
.tk-month-analysis-card.excluded { border-color:rgba(251,191,36,.95); background:rgba(255,251,235,.95); }
.tk-month-analysis-card-head {
  display:flex; align-items:baseline; justify-content:space-between; gap:12px; flex-wrap:wrap;
  border-bottom:1px solid rgba(148,163,184,.35); padding-bottom:8px; margin-bottom:2px;
}
.tk-month-analysis-card-head span {
  color:#64748b; font-size:.82rem; font-weight:800; text-transform:uppercase; letter-spacing:.04em;
}
.tk-month-analysis-card-head strong { color:#0f172a; font-size:1.25rem; letter-spacing:-0.03em; }
.tk-month-analysis-epic-list {
  max-height:280px; overflow:auto; margin:0 -4px; padding:4px;
  display:flex; flex-direction:column; gap:10px;
}
.tk-month-analysis-epic-row {
  padding:8px 10px; border-radius:10px; background:rgba(255,255,255,.72); border:1px solid rgba(203,213,225,.65);
  font-size:.88rem; line-height:1.35;
}
.tk-month-analysis-card.excluded .tk-month-analysis-epic-row { background:rgba(255,255,255,.85); }
.tk-month-epic-heading { display:flex; flex-wrap:wrap; align-items:baseline; gap:8px 10px; margin-bottom:4px; }
.tk-month-epic-link {
  font-weight:800; color:#1d4ed8; text-decoration:none; flex-shrink:0;
}
.tk-month-epic-link:hover { text-decoration:underline; }
.tk-month-analysis-epic-title { color:#334155; font-weight:600; }
.tk-month-analysis-exclude-reason {
  margin-top:6px; padding-top:6px; border-top:1px dashed rgba(251,191,36,.55);
  color:#92400e; font-size:.82rem; font-weight:600; line-height:1.45;
}
.tk-month-analysis-planner-remark {
  margin-top:6px; font-size:.8rem; color:#57534e; line-height:1.4;
}
.tk-month-analysis-planner-remark .lbl { font-weight:700; color:#78716c; }
.tk-month-analysis-list-empty { color:#94a3b8; font-size:.86rem; font-style:italic; padding:6px 4px; }
.tk-month-analysis-scope-reconcile { margin-top:10px; max-width:900px; }
.footnote { color:var(--muted); line-height:1.5; font-size:.92rem; }

/* ── Hierarchy table ──────────────────────────────────────────────── */
.table-frame { width:100%; overflow:auto; }
.epic-table { width:max(100%,1400px); border-collapse:collapse; table-layout:fixed; background:#fff; }
.epic-table th, .epic-table td { padding:10px 12px; border:1px solid #d7e0ea; text-align:left; vertical-align:top; }
.epic-table thead th {
  position:sticky; top:0; z-index:18;
  background:linear-gradient(180deg,#d4e2f0,#c0d2e6);
  border-bottom-color:#a6bcd0; text-transform:uppercase;
  font-size:.78rem; letter-spacing:.05em; color:#3d5673;
}
.row-toggle {
  appearance:none; width:28px; height:28px; border:1px solid rgba(192,206,218,.90);
  border-radius:7px; background:#fff; color:#315b8a; font-size:1rem; font-weight:800;
  line-height:1; cursor:pointer;
}
.row-toggle:hover { background:#f0f5fb; }
.row-toggle[disabled] { opacity:.35; cursor:default; }
.epic-row, .epic-row td { background:var(--row-epic); }
.epic-row:hover, .epic-row:hover td { background:var(--row-epic-hover); }
.story-row, .story-row td { background:var(--row-story); }
.story-row:hover, .story-row:hover td { background:var(--row-story-hover); }
.subtask-row, .subtask-row td { background:var(--row-subtask); }
.subtask-row:hover, .subtask-row:hover td { background:var(--row-subtask-hover); }
.bug-row, .bug-row td { background:var(--row-bug); }
.bug-row:hover, .bug-row:hover td { background:var(--row-bug-hover); }
.epic-title-text { font-size:.98rem; font-weight:800; line-height:1.3; color:#102033; }
.story-title-text { font-size:.92rem; font-weight:600; padding-left:22px; }
.subtask-title-text { font-size:.88rem; padding-left:44px; color:#37475a; }
.row-toggle-spacer {
  display:inline-block; width:28px; height:28px; vertical-align:middle;
}
.subtask-worklog-panel > td.subtask-worklog-panel-cell {
  background:linear-gradient(180deg,#e8f4fc,#f0f9ff);
  padding:12px 12px 14px 56px;
  border-left:3px solid #7dd3fc;
  vertical-align:top;
}
.worklog-nested-table {
  width:100%;
  max-width:920px;
  border-collapse:collapse;
  font-size:.84rem;
  background:#fff;
  border-radius:8px;
  overflow:hidden;
  box-shadow:0 1px 3px rgba(15,23,42,.08);
}
.worklog-nested-table th {
  background:linear-gradient(180deg,#bfdbfe,#93c5fd);
  text-transform:uppercase;
  font-size:.72rem;
  letter-spacing:.04em;
  padding:8px 10px;
  text-align:left;
  color:#1e3a5f;
}
.worklog-nested-table td { padding:8px 10px; border-top:1px solid #e2e8f0; color:#334155; }
.worklog-nested-table tbody tr:hover td { background:#f8fafc; }
.jira-link {
  display:inline-flex; align-items:center; justify-content:center;
  min-width:34px; height:34px; padding:0 6px; border:1px solid rgba(29,78,216,.22);
  border-radius:9px; background:#fff; color:#1d4ed8; text-decoration:none;
  font-size:.75rem; font-weight:800; transition:all .15s;
}
.jira-link:hover { border-color:rgba(29,78,216,.55); background:#eff6ff; }

/* ── Table legend ─────────────────────────────────────────────────── */
.table-legend {
  display:flex; gap:10px; flex-wrap:wrap; margin:12px 0;
  padding:12px 14px; border:1px solid rgba(200,214,228,.75);
  border-radius:var(--radius-sm); background:rgba(248,250,255,.90);
}
.table-legend-item {
  display:inline-flex; align-items:center; gap:8px; padding:6px 10px;
  border-radius:999px; background:#fff; border:1px solid rgba(210,222,232,.90);
  font-size:.84rem; font-weight:700;
}
.table-legend-swatch { width:12px; height:12px; border-radius:999px; border:1px solid rgba(16,32,51,.12); }

/* ── Gantt view ───────────────────────────────────────────────────── */
.gantt-product-grid { display:grid; gap:16px; }
.gantt-product-section {
  border:1px solid rgba(200,214,228,.75); border-radius:var(--radius-md);
  padding:16px; background:rgba(252,253,255,.96); box-shadow:var(--shadow-sm);
}
.gantt-product-section[data-hidden="true"] { display:none; }
.gantt-label { font-size:12px; font-weight:800; fill:#203141; }
.gantt-meta { font-size:11px; fill:#627487; }

/* ── Drawer ───────────────────────────────────────────────────────── */
.drawer-overlay { position:fixed; inset:0; background:rgba(15,23,42,.45); z-index:99; display:none; }
.drawer-overlay.open { display:flex; justify-content:flex-end; }
.drawer-panel {
  width:min(860px,94vw); height:100vh; background:#fff;
  box-shadow:-12px 0 48px rgba(15,23,42,.35); overflow-y:auto; padding:24px;
}
.drawer-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; }
.drawer-close {
  appearance:none; border:1px solid rgba(192,206,218,.90); background:#fff; color:#203141;
  border-radius:999px; padding:8px 14px; font-size:.84rem; font-weight:700; cursor:pointer;
}
.drawer-epics-list { display:grid; gap:10px; }
.drawer-epic-card {
  padding:12px 14px; border:1px solid var(--line); border-radius:var(--radius-sm);
  background:var(--panel-soft);
}
.drawer-epic-card strong { color:#102033; }
.drawer-epic-card .meta { color:var(--muted); font-size:.88rem; }
.drawer-prose { font-size:.9rem; line-height:1.45; color:var(--muted); margin:0 0 12px; }
.drawer-prose strong { color:#102033; }
.drawer-employee-section { margin:20px 0; }
.drawer-employee-section h3 {
  font-size:1.02rem; font-weight:800; margin:0 0 6px; color:#102033;
  letter-spacing:-.02em; border-bottom:1px solid var(--line); padding-bottom:6px;
}
.drawer-emp-table { width:100%; border-collapse:collapse; font-size:.86rem; margin-top:6px; }
.drawer-emp-table th, .drawer-emp-table td { border:1px solid var(--line); padding:7px 10px; text-align:left; vertical-align:top; }
.drawer-emp-table th { background:var(--panel-soft); color:#1e2a3b; }
.drawer-emp-table th.drawer-emp-idx, .drawer-emp-table td.drawer-emp-idx { width:2.5em; text-align:right; color:#3d4a5c; }
.drawer-empty { font-size:.88rem; color:var(--muted); margin:4px 0; font-style:italic; }
.capacity-field-value--clickable {
  cursor:pointer; border-bottom:1px dotted rgba(29,78,216,.45); color:#1d4ed8;
  transition: color .12s, background .12s, border-color .12s; border-radius:4px; padding:2px 4px; margin:-2px -4px;
}
.capacity-field-value--clickable:hover { color:#0f3d9e; background:rgba(37,99,235,.08); }

/* ── Misc ─────────────────────────────────────────────────────────── */
.view-section[hidden] { display:none; }
.empty-state { padding:16px; border:1px dashed rgba(192,210,226,.80); border-radius:var(--radius-sm); color:var(--muted); background:rgba(251,253,255,.90); text-align:center; }
a { color:#1d4ed8; }

/* --- RMI Estimation & Scheduling Table (IPP reference layout) --- */
.rmi-schedule-panel { margin-bottom:var(--gutter); }
.rmi-schedule-header-bar {
  display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-bottom:12px;
}
.rmi-schedule-header-bar h2 { margin:0; }
.rmi-schedule-controls { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.rmi-schedule-year-label { font-size:.82rem; font-weight:600; color:var(--muted); }
.rmi-schedule-year-select {
  font-size:.82rem; padding:4px 10px; border:1px solid var(--line); border-radius:6px; background:var(--panel); cursor:pointer;
}
.rmi-schedule-panel .table-frame { overflow-x:auto; overflow-y:auto; max-width:100%; max-height:80vh; }
.rmi-schedule-table {
  border-collapse:separate; border-spacing:0; font-size:.78rem; min-width:1200px;
}
.rmi-schedule-table th, .rmi-schedule-table td {
  padding:6px 8px; border:1px solid var(--line); text-align:center; white-space:nowrap;
  overflow:hidden; text-overflow:ellipsis;
}
.rmi-schedule-table th { font-weight:700; }
.rmi-schedule-table thead th { position:sticky; z-index:2; background:#f1f5f9; }
.rmi-sched-header-groups th { top:0; background:var(--panel-soft) !important; }
.rmi-sched-header-cols th { top:var(--rmi-sched-row1-h,31px); }
.rmi-schedule-table th:nth-child(-n+6), .rmi-schedule-table td:nth-child(-n+6) { position:sticky; z-index:2; }
.rmi-schedule-table thead th:nth-child(-n+6) { z-index:4; }
.rmi-schedule-table th:nth-child(1), .rmi-schedule-table td:nth-child(1) {
  left:0; min-width:36px; width:36px; background:var(--panel); text-align:center; color:var(--muted); font-size:.7rem;
}
.rmi-schedule-table th:nth-child(2), .rmi-schedule-table td:nth-child(2) {
  left:36px; min-width:360px; width:360px; white-space:normal; word-break:break-word; background:var(--panel);
}
.rmi-schedule-table th:nth-child(3), .rmi-schedule-table td:nth-child(3) { left:396px; min-width:90px; width:90px; background:var(--panel); }
.rmi-schedule-table th:nth-child(4), .rmi-schedule-table td:nth-child(4) { left:486px; min-width:80px; width:80px; background:var(--panel); }
.rmi-schedule-table th:nth-child(5), .rmi-schedule-table td:nth-child(5) { left:566px; min-width:70px; width:70px; background:var(--panel); }
.rmi-schedule-table th:nth-child(6), .rmi-schedule-table td:nth-child(6) {
  left:636px; min-width:70px; width:70px; background:var(--panel); border-right:2px solid #94a3b8;
}
.rmi-sched-header-cols th { background:#f1f5f9 !important; font-size:.74rem; text-transform:uppercase; letter-spacing:.04em; }
.rmi-sched-header-groups th { background:var(--panel-soft) !important; border-bottom:none; }
.rmi-sched-header-groups th.rmi-sched-group-estimation { background:#1e3a5f !important; color:#fff !important; }
.rmi-sched-header-groups th.rmi-sched-group-scheduling { background:#d97706 !important; color:#fff !important; }
.rmi-sched-col-rmi { text-align:left !important; }
.rmi-sched-month { min-width:62px; width:62px; }
.rmi-schedule-table td.rmi-sched-cell-rmi { text-align:left; font-weight:600; }
.rmi-sched-jira-link {
  display:inline-flex; align-items:center; justify-content:center; width:18px; height:18px; margin-left:4px;
  vertical-align:middle; border-radius:3px; background:#0052cc; color:#fff; font-size:10px; font-weight:800;
  text-decoration:none; line-height:1;
}
.rmi-sched-jira-link:hover { background:#0747a6; }
.rmi-sched-product-group td {
  background:#f8fafc !important; font-weight:800; font-size:.76rem; text-transform:uppercase; letter-spacing:.06em; border-top:2px solid var(--line);
}
.rmi-sched-product-group td:first-child { border-left:4px solid var(--muted); }
.rmi-sched-product-group td.rmi-sched-group-label { text-align:left; }
.rmi-sched-product-subtotal td {
  background:#f1f5f9 !important; font-weight:700; font-size:.76rem; border-top:1px solid var(--line);
}
.rmi-sched-product-subtotal td:first-child { border-left:4px solid var(--muted); text-align:right; padding-right:12px; }
.rmi-sched-grand-total td {
  background:#e2e8f0 !important; font-weight:800; font-size:.78rem; border-top:2px solid #94a3b8;
}
.rmi-sched-grand-total td:first-child { text-align:right; padding-right:12px; }
.rmi-sched-epic-row:nth-child(even) td { background:#fafbfd !important; }
.rmi-sched-epic-row:hover td { background:#eef2ff !important; }
.rmi-sched-status-pill {
  display:inline-block; padding:2px 8px; border-radius:10px; font-size:.68rem; font-weight:700; text-transform:capitalize;
  background:#e2e8f0; color:#334155;
}
.rmi-sched-status-pill[data-status-lower="done"], .rmi-sched-status-pill[data-status-lower="delivered"] {
  background:#d1fae5; color:#065f46;
}
.rmi-sched-status-pill[data-status-lower="in progress"], .rmi-sched-status-pill[data-status-lower="in-progress"] {
  background:#dbeafe; color:#1e40af;
}
.rmi-sched-status-pill[data-status-lower="to do"], .rmi-sched-status-pill[data-status-lower="to-do"] {
  background:#fef3c7; color:#92400e;
}
.rmi-sched-status-pill[data-status-lower="on hold"], .rmi-sched-status-pill[data-status-lower="on-hold"] {
  background:#dbeafe; color:#1e40af;
}
.rmi-sched-product-cards {
  display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:10px; margin-bottom:14px;
}
.rmi-sched-pcard {
  display:flex; flex-direction:column; gap:2px; padding:12px 16px; border:2px solid var(--line); border-radius:12px;
  background:var(--panel); cursor:pointer; transition:border-color .15s,box-shadow .15s; user-select:none;
  border-left:4px solid var(--product-accent,var(--line));
}
.rmi-sched-pcard:hover { box-shadow:0 2px 8px rgba(0,0,0,.06); }
.rmi-sched-pcard:focus-visible { outline:2px solid var(--product-accent,#2563eb); outline-offset:2px; }
.rmi-sched-pcard.active { border-color:var(--product-accent,#102033); box-shadow:0 2px 12px rgba(0,0,0,.10); }
.rmi-sched-pcard-label { font-size:.78rem; font-weight:700; color:var(--text); }
.rmi-sched-pcard-value { font-size:1.3rem; font-weight:800; color:var(--product-accent,var(--text)); font-style:italic; }
.rmi-sched-pcard-meta { font-size:.68rem; color:var(--muted); }

/* ── Responsive ───────────────────────────────────────────────────── */
@media (max-width:900px) {
  .page { padding:24px 12px 40px; }
  .estimate-cards-group { grid-template-columns:repeat(2,1fr); }
  .metric-card-hero { grid-column:auto; }
  .epic-table thead th { position:static; }
  .capacity-calculator { grid-template-columns:1fr; }
  .capacity-calc-title {
    writing-mode:horizontal-tb; transform:none; padding:10px 16px;
    border-radius:var(--radius-lg) var(--radius-lg) 0 0;
  }
  .capacity-calc-body { border-radius:0 0 var(--radius-lg) var(--radius-lg); }
  .capacity-results-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
"""

# ── Rich Report JS ───────────────────────────────────────────────────

_REPORT_JS = """
(function () {
  "use strict";

  /* ── Data ─────────────────────────────────────────────────────────── */
  var DATA = JSON.parse(document.getElementById("rmi-report-data").textContent);
  var epics = DATA.epics || [];
  var capacitySource = DATA.capacity_source || {};

  /* ── State ────────────────────────────────────────────────────────── */
  var activeProduct = "all";
  var unit = "hours";
  var SEC_PER_HOUR = 3600;
  var SEC_PER_DAY  = 28800;

  /* ── Utils ────────────────────────────────────────────────────────── */
  function fmtSec(v) {
    var n = Number(v || 0);
    var prefix = n < 0 ? "-" : "";
    var abs = Math.abs(n);
    if (unit === "days") return prefix + (abs / SEC_PER_DAY).toLocaleString(undefined, {maximumFractionDigits:2}) + " d";
    return prefix + (abs / SEC_PER_HOUR).toLocaleString(undefined, {maximumFractionDigits:2}) + " h";
  }

  function searchQuery() {
    var el = document.getElementById("epic-search");
    return el ? el.value.trim().toLowerCase() : "";
  }

  function scopedEpics() {
    var q = searchQuery();
    return epics.filter(function (e) {
      return (activeProduct === "all" || e.product === activeProduct) &&
             (!q || (e.jira_id + " " + e.title).toLowerCase().indexOf(q) >= 0);
    });
  }

  function scopedTotals() {
    var keys = ["most_likely_seconds","optimistic_seconds","pessimistic_seconds","calculated_seconds",
                "tk_approved_seconds","jira_original_estimate_seconds","story_estimate_seconds",
                "subtask_estimate_seconds","logged_seconds"];
    var t = {epic_count: 0};
    keys.forEach(function (k) { t[k] = 0; });
    scopedEpics().forEach(function (e) {
      t.epic_count++;
      keys.forEach(function (k) { t[k] += Number(e[k] || 0); });
    });
    return t;
  }

  /* ── Duration value update ────────────────────────────────────────── */
  function updateAllDurations() {
    document.querySelectorAll(".duration-value").forEach(function (el) {
      if (el.closest(".rmi-sched-pcard")) return;
      var sec = Number(el.dataset.seconds || 0);
      el.textContent = fmtSec(sec);
    });
  }

  /* ── Metric cards ─────────────────────────────────────────────────── */
  var METRIC_DEFS = [
    {key:"epic_count",       label:"Total # of RMI Epics", cls:"metric-card-teal",  type:"count",    meta:"Epic parents in the selected product scope"},
    {key:"optimistic_seconds",  label:"Optimistic",  cls:"metric-estimate-step-1", type:"duration", meta:"Workbook optimistic total for the selected epic set", isEstimate:true},
    {key:"most_likely_seconds", label:"Most Likely", cls:"metric-estimate-step-2", type:"duration", meta:"Workbook most likely total for the selected epic set", isEstimate:true},
    {key:"pessimistic_seconds", label:"Pessimistic", cls:"metric-estimate-step-3", type:"duration", meta:"Workbook pessimistic total for the selected epic set", isEstimate:true},
    {key:"calculated_seconds",  label:"Calculated",  cls:"metric-estimate-step-4", type:"duration", meta:"Workbook calculated estimate total for the selected epic set", isEstimate:true},
    {key:"tk_approved_seconds",  label:"TK Approved", cls:"metric-card-emerald", type:"duration", meta:"TK approved total for the selected epic set", hero:true},
    {key:"jira_original_estimate_seconds", label:"Epic Estimates",    cls:"metric-card-indigo", type:"duration", meta:"Epic-level Jira original estimate total", diagnosticsHide:true},
    {key:"story_estimate_seconds",         label:"Story Estimates",   cls:"metric-card-slate",  type:"duration", meta:"Story-level original estimate total (excludes subtasks)", diagnosticsHide:true},
    {key:"subtask_estimate_seconds",       label:"Subtask Estimates", cls:"metric-card-violet", type:"duration", meta:"Subtask-level original estimate total", diagnosticsHide:true},
    {key:"logged_seconds",                 label:"Logged",            cls:"metric-card-rose",   type:"duration", meta:"Total hours logged across all stories/subtasks"}
  ];

  function diagnosticsEnabled() {
    var toggle = document.getElementById("diagnostics-toggle-enabled");
    return Boolean(toggle && toggle.checked);
  }

  function renderMetrics() {
    var t = scopedTotals();
    var grid = document.getElementById("metric-grid");
    if (!grid) return;
    var diagnosticsMode = diagnosticsEnabled();

    var estimateDefs = METRIC_DEFS.filter(function (d) { return d.isEstimate; });
    var otherDefs    = METRIC_DEFS.filter(function (d) {
      if (d.isEstimate) return false;
      if (!diagnosticsMode && d.diagnosticsHide) return false;
      return true;
    });

    var html = "";
    /* epic count */
    var ec = otherDefs.find(function (d) { return d.key === "epic_count"; });
    if (ec) {
      html += '<section class="metric-card ' + ec.cls + '" data-metric-key="' + ec.key + '">'
            + '<div class="metric-label">' + ec.label + '</div>'
            + '<div class="metric-value-wrap"><div class="metric-value">' + t[ec.key] + '</div></div>'
            + '<div class="metric-meta">' + ec.meta + '</div></section>';
    }
    /* estimate range group */
    html += '<div class="estimate-cards-group">';
    estimateDefs.forEach(function (d) {
      var v = t[d.key] || 0;
      html += '<section class="metric-card ' + d.cls + '" data-metric-key="' + d.key + '">'
            + '<div class="metric-label">' + d.label + '</div>'
            + '<div class="metric-value-wrap"><span class="metric-value duration-value" data-seconds="' + v + '">' + fmtSec(v) + '</span></div>'
            + '<div class="metric-meta">' + d.meta + '</div></section>';
    });
    html += '</div>';

    /* remaining cards */
    otherDefs.filter(function (d) { return d.key !== "epic_count"; }).forEach(function (d) {
      var v = t[d.key] || 0;
      var heroClass = d.hero ? " metric-card-hero metric-card-clickable" : " metric-card-clickable";
      var clickIcon = d.hero ? "" : '<span class="metric-click-icon" aria-hidden="true">&#x2192;</span>';
      html += '<section class="metric-card ' + d.cls + heroClass + '" data-metric-key="' + d.key + '" role="button" tabindex="0">'
            + '<div class="metric-label">' + d.label + clickIcon + '</div>'
            + '<div class="metric-value-wrap"><span class="metric-value duration-value" data-seconds="' + v + '">' + fmtSec(v) + '</span></div>'
            + '<div class="metric-meta">' + d.meta + '. Click to view contributing epics.' + '</div></section>';
    });

    grid.innerHTML = html;
    renderProductSummary();
    renderCapacity();
  }

  /* ── Product summary cards ────────────────────────────────────────── */
  var PRODUCT_ACCENTS = {
    "Digital Log": "#7c3aed",
    "Fintech Fuel": "#b45309",
    "OmniChat": "#2563eb",
    "OmniConnect": "#0f766e"
  };
  var PRODUCT_COLORS_FALLBACK = ["#1d4ed8","#be123c","#6d28d9","#047857","#4338ca","#334155"];
  /* Gantt row colors (must exist — renderGantt runs during applyFilters before schedule init) */
  var PRODUCT_COLORS = PRODUCT_COLORS_FALLBACK;

  function renderProductSummary() {
    var grid = document.getElementById("product-summary-grid");
    if (!grid) return;
    var products = [];
    var seen = {};
    epics.forEach(function (e) { var p = e.product || "Unassigned"; if (!seen[p]) { seen[p] = true; products.push(p); } });
    products.sort();

    var allTk = epics.reduce(function (s, e) { return s + Number(e.tk_approved_seconds || 0); }, 0);
    var allCount = epics.length;
    var html = '<section class="product-summary-card' + (activeProduct === "all" ? " active" : "") + '" data-product-summary="all" style="--product-accent:#102033">'
             + '<div class="product-summary-label">All Products</div>'
             + '<div class="product-summary-value"><span class="product-summary-duration duration-value" data-seconds="' + allTk + '">' + fmtSec(allTk) + '</span></div>'
             + '<div class="product-summary-meta">' + allCount + ' RMIs/Epics &nbsp;&bull;&nbsp; Total TK Approved</div></section>';

    var fallbackIdx = 0;
    products.forEach(function (p) {
      var color = PRODUCT_ACCENTS[p] || PRODUCT_COLORS_FALLBACK[(fallbackIdx++) % PRODUCT_COLORS_FALLBACK.length];
      var productEpics = epics.filter(function (e) { return e.product === p; });
      var tk = productEpics.reduce(function (s, e) { return s + Number(e.tk_approved_seconds || 0); }, 0);
      var count = productEpics.length;
      html += '<section class="product-summary-card' + (activeProduct === p ? " active" : "") + '" data-product-summary="' + p + '" style="--product-accent:' + color + '">'
            + '<div class="product-summary-label">' + p + '</div>'
            + '<div class="product-summary-value"><span class="product-summary-duration duration-value" data-seconds="' + tk + '">' + fmtSec(tk) + '</span></div>'
            + '<div class="product-summary-meta">' + count + ' RMIs/Epics &nbsp;&bull;&nbsp; Total TK Approved</div></section>';
    });
    grid.innerHTML = html;

    grid.querySelectorAll(".product-summary-card").forEach(function (card) {
      card.addEventListener("click", function () {
        activeProduct = card.dataset.productSummary;
        applyFilters();
      });
    });
  }

  /* ── Capacity (auto-sourced from RLT + teams) ────────────────────── */
  var rltEmployees = capacitySource.rlt_employees || [];
  var teams = capacitySource.teams || [];
  var memberRecords = capacitySource.member_records || {};
  var capacityProfiles = capacitySource.capacity_profiles || [];
  var capacityAllMembers = new Set();
  var selectedMembers = new Set();
  var teamMemberMap = Object.create(null);

  function parseIsoDateToLocal(s) {
    s = String(s || "").trim();
    if (s.length < 10) return null;
    var y = parseInt(s.slice(0, 4), 10);
    var m = parseInt(s.slice(5, 7), 10) - 1;
    var d = parseInt(s.slice(8, 10), 10);
    if (!y || m < 0) return null;
    return new Date(y, m, d);
  }

  function findCapacityProfileForMonth(month) {
    if (!month || !capacityProfiles.length) return null;
    var p = String(month).split("-");
    if (p.length < 2) return null;
    var y = parseInt(p[0], 10), mo = parseInt(p[1], 10);
    if (!y || !mo) return null;
    var monthStart = new Date(y, mo - 1, 1);
    var monthEnd = new Date(y, mo, 0);
    var mst = monthStart.getTime();
    var met = monthEnd.getTime();
    for (var i = 0; i < capacityProfiles.length; i++) {
      var pr = capacityProfiles[i];
      var fs = parseIsoDateToLocal(pr && pr.from_date);
      var fe = parseIsoDateToLocal(pr && pr.to_date);
      if (!fs || !fe) continue;
      if (fe.getTime() < mst || fs.getTime() > met) continue;
      return pr;
    }
    return null;
  }

  function isGlobalCapacityScope() {
    return isAllMembersSelected() && activeProduct === "all";
  }

  /* Searchable multi-select team filter */
  (function initTeamSelector() {
    function escAttr(s) {
      return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    }
    var root = document.getElementById("capacity-team-ms");
    var trigger = document.getElementById("capacity-team-ms-trigger");
    var panel = document.getElementById("capacity-team-ms-panel");
    var list = document.getElementById("capacity-team-ms-list");
    var search = document.getElementById("capacity-team-ms-search");
    var labelEl = document.getElementById("capacity-team-ms-label");
    var allCb = document.getElementById("capacity-team-ms-all");
    var allText = document.getElementById("capacity-team-ms-all-text");
    var clearBtn = document.getElementById("capacity-team-ms-clear");
    if (!root || !trigger || !panel || !list || !labelEl) return;

    function listFromSet(setObj) {
      var out = [];
      setObj.forEach(function (v) { out.push(v); });
      out.sort(function (a, b) { return String(a).localeCompare(String(b)); });
      return out;
    }

    function isAllSelectedNow() {
      return capacityAllMembers.size > 0 && selectedMembers.size === capacityAllMembers.size;
    }

    function selectAllMembers() {
      selectedMembers.clear();
      capacityAllMembers.forEach(function (m) { selectedMembers.add(m); });
    }

    function selectDefaultMembers() {
      selectedMembers.clear();
      capacityAllMembers.forEach(function (m) {
        var rec = memberRecords[m] || {};
        if (!Boolean(rec.resigned)) selectedMembers.add(m);
      });
    }

    function clearAllMembers() {
      selectedMembers.clear();
    }

    function membersForTeam(teamName) {
      return teamMemberMap[teamName] || [];
    }

    function memberStatusLabel(name) {
      var rec = memberRecords[name] || {};
      return Boolean(rec.resigned) ? "Resigned" : "Active";
    }

    function refreshCheckboxStates() {
      list.querySelectorAll(".capacity-ms-member-cb").forEach(function (cb) {
        cb.checked = selectedMembers.has(cb.value);
      });
      list.querySelectorAll(".capacity-ms-team-cb").forEach(function (cb) {
        var tn = cb.getAttribute("data-team-name") || "";
        var members = membersForTeam(tn);
        var selectedCount = 0;
        members.forEach(function (m) {
          if (selectedMembers.has(m)) selectedCount += 1;
        });
        cb.indeterminate = selectedCount > 0 && selectedCount < members.length;
        cb.checked = members.length > 0 && selectedCount === members.length;
      });
      if (allCb) {
        allCb.indeterminate = selectedMembers.size > 0 && !isAllSelectedNow();
        allCb.checked = isAllSelectedNow();
      }
    }

    (function setAllEmpLabel() {
      var msel = document.getElementById("tk-month-select");
      var m = msel ? msel.value : "";
      var p = findCapacityProfileForMonth(m);
      var n = p ? Math.max(0, Math.round(Number(p.employee_count) || 0)) : rltEmployees.length;
      if (allText) allText.textContent = "Select all (" + n + ")";
    })();

    teams.forEach(function (t) {
      var teamName = String(t.team_name || "").trim();
      if (!teamName) return;
      var members = [];
      var seen = Object.create(null);
      (t.assignees || []).forEach(function (a) {
        var mn = String(a || "").trim();
        if (!mn || seen[mn]) return;
        seen[mn] = true;
        members.push(mn);
      });
      members.sort(function (a, b) { return a.localeCompare(b); });
      teamMemberMap[teamName] = members;
      members.forEach(function (m) { capacityAllMembers.add(m); });

      var group = document.createElement("div");
      group.className = "capacity-ms-team-group";
      group.setAttribute("data-team-name", teamName);
      group.setAttribute("data-search", (teamName + " " + members.join(" ")).toLowerCase());
      var teamRow = document.createElement("label");
      teamRow.className = "capacity-ms-item capacity-ms-team-row";
      teamRow.innerHTML = '<input type="checkbox" class="capacity-ms-team-cb" data-team-name="' + escAttr(teamName) + '" />'
        + '<span class="capacity-ms-item-text">' + escAttr(teamName) + " (" + members.length + ")</span>";
      group.appendChild(teamRow);
      var membersWrap = document.createElement("div");
      membersWrap.className = "capacity-ms-members";
      members.forEach(function (m) {
        var memberRow = document.createElement("label");
        memberRow.className = "capacity-ms-item capacity-ms-member-row";
        memberRow.setAttribute("data-team-name", teamName);
        memberRow.setAttribute("data-member-name", m);
        var statusLabel = memberStatusLabel(m);
        memberRow.innerHTML = '<input type="checkbox" class="capacity-ms-member-cb" value="' + escAttr(m) + '" data-team-name="' + escAttr(teamName) + '" />'
          + '<span class="capacity-ms-item-text">' + escAttr(m) + " (" + escAttr(statusLabel) + ")</span>";
        membersWrap.appendChild(memberRow);
      });
      group.appendChild(membersWrap);
      list.appendChild(group);
    });
    selectDefaultMembers();

    function updateTriggerLabel() {
      if (isAllSelectedNow()) {
        labelEl.textContent = "All Employees";
        return;
      }
      if (selectedMembers.size === 0) {
        labelEl.textContent = "No members selected";
        return;
      }
      if (selectedMembers.size === 1) {
        var one = listFromSet(selectedMembers);
        labelEl.textContent = one.length ? one[0] : "1 member selected";
        return;
      }
      labelEl.textContent = selectedMembers.size + " members selected";
    }

    function applySearchFilter() {
      var q = (search && search.value ? search.value : "").trim().toLowerCase();
      var allRow = list.querySelector(".capacity-ms-item-all");
      if (allRow) {
        var hitAll = !q || "select all".indexOf(q) >= 0 || "all".indexOf(q) === 0;
        allRow.classList.toggle("capacity-ms-hidden", !hitAll);
      }
      list.querySelectorAll(".capacity-ms-team-group").forEach(function (group) {
        var haystack = String(group.getAttribute("data-search") || "");
        group.classList.toggle("capacity-ms-hidden", q.length > 0 && haystack.indexOf(q) < 0);
      });
    }

    function setOpen(open) {
      panel.hidden = !open;
      trigger.setAttribute("aria-expanded", open ? "true" : "false");
      root.classList.toggle("capacity-ms-open", open);
      var calcCard = root.closest(".capacity-calculator");
      if (calcCard) calcCard.classList.toggle("capacity-calculator--dropdown-open", open);
      if (!open && search) {
        search.value = "";
        applySearchFilter();
      }
      if (open && search) {
        try { search.focus(); } catch (e) {}
      }
    }

    function onDocClick(e) {
      if (!root.contains(e.target)) setOpen(false);
    }

    trigger.addEventListener("click", function (e) {
      e.stopPropagation();
      setOpen(panel.hidden);
    });
    document.addEventListener("click", onDocClick);
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && !panel.hidden) setOpen(false);
    });
    if (search) {
      search.addEventListener("click", function (e) { e.stopPropagation(); });
      search.addEventListener("input", applySearchFilter);
    }

    list.addEventListener("change", function (e) {
      var t = e.target;
      if (!t || t.type !== "checkbox") return;
      if (t.classList.contains("capacity-ms-all")) {
        if (t.checked) selectAllMembers();
        else clearAllMembers();
      } else if (t.classList.contains("capacity-ms-team-cb")) {
        var teamName = t.getAttribute("data-team-name") || "";
        membersForTeam(teamName).forEach(function (m) {
          if (t.checked) selectedMembers.add(m);
          else selectedMembers.delete(m);
        });
      } else if (t.classList.contains("capacity-ms-member-cb")) {
        if (t.checked) selectedMembers.add(t.value);
        else selectedMembers.delete(t.value);
      }
      refreshCheckboxStates();
      updateTriggerLabel();
      renderCapacity();
      renderMetrics();
    });

    if (clearBtn) {
      clearBtn.addEventListener("click", function (e) {
        e.stopPropagation();
        clearAllMembers();
        refreshCheckboxStates();
        updateTriggerLabel();
        renderCapacity();
        renderMetrics();
      });
    }

    refreshCheckboxStates();
    updateTriggerLabel();
  })();

  function isAllMembersSelected() {
    return capacityAllMembers.size > 0 && selectedMembers.size === capacityAllMembers.size;
  }

  function selectedTeamNames() {
    var out = [];
    teams.forEach(function (t) {
      var tn = String((t && t.team_name) || "");
      if (!tn) return;
      var members = teamMemberMap[tn] || [];
      var hasSelected = false;
      for (var i = 0; i < members.length; i++) {
        if (selectedMembers.has(members[i])) {
          hasSelected = true;
          break;
        }
      }
      if (hasSelected) out.push(tn);
    });
    out.sort();
    return out;
  }

  function scopedCapacityAssignees() {
    /* Member selection is the source of truth; members are deduped globally. */
    if (!isAllMembersSelected()) {
      var out = [];
      selectedMembers.forEach(function (m) { out.push(m); });
      out.sort();
      return out;
    }
    if (activeProduct !== "all") {
      var byProduct = capacitySource.employees_by_product || {};
      var productAssignees = byProduct[activeProduct] || [];
      if (productAssignees.length) return productAssignees;
    }
    return rltEmployees.length ? rltEmployees : (capacitySource.employees_by_product || {}).all || [];
  }

  function scopedLeaveHours(month) {
    var byMonth = capacitySource.leaves_hours_by_month || {};
    var byMonthAssignee = capacitySource.leaves_hours_by_month_assignee || {};
    if (isGlobalCapacityScope()) {
      return Number(byMonth[month] || 0);
    }
    var assignees = scopedCapacityAssignees();
    if (!assignees.length) return Number(byMonth[month] || 0);
    var ma = byMonthAssignee[month] || {};
    var total = 0;
    assignees.forEach(function (a) { total += Number(ma[a] || 0); });
    return total;
  }

  function monthToggleState() {
    var started = document.getElementById("tk-start-month-enabled") && document.getElementById("tk-start-month-enabled").checked;
    var delivered = document.getElementById("tk-month-enabled") && document.getElementById("tk-month-enabled").checked;
    var through = document.getElementById("tk-through-month-enabled") && document.getElementById("tk-through-month-enabled").checked;
    return { started: started, delivered: delivered, through: through, anyToggle: started || delivered || through };
  }

  function renderCapacity() {
    var monthEl = document.getElementById("tk-month-select");
    var month = monthEl ? monthEl.value : "";
    var prof = findCapacityProfileForMonth(month);
    var global = isGlobalCapacityScope();
    var assignees = scopedCapacityAssignees();
    var employees = global
      ? (prof ? Math.max(0, Math.round(Number(prof.employee_count) || 0)) : assignees.length)
      : assignees.length;
    var leaveHours = scopedLeaveHours(month);
    var leaveSec = leaveHours * SEC_PER_HOUR;

    var allText = document.getElementById("capacity-team-ms-all-text");
    if (allText) {
      var nLab = prof ? Math.max(0, Math.round(Number(prof.employee_count) || 0)) : rltEmployees.length;
      allText.textContent = "Select all (" + nLab + ")";
    }

    var empEl = document.getElementById("capacity-employees-val");
    var leaveEl = document.getElementById("capacity-leaves-val");
    if (empEl) empEl.textContent = String(employees);
    if (leaveEl) leaveEl.textContent = fmtSec(leaveSec);

    var parts = month.split("-").map(Number);
    var year = parts[0], mon = parts[1];
    var workingDays = 0;
    if (year && mon) {
      var d = new Date(year, mon - 1, 1);
      while (d.getMonth() === mon - 1) { var day = d.getDay(); if (day !== 0 && day !== 6) workingDays++; d.setDate(d.getDate() + 1); }
    }
    var capacity = employees * workingDays * SEC_PER_DAY;
    var availability = capacity - leaveSec;

    var capEl = document.getElementById("capacity-value");
    var availEl = document.getElementById("availability-value");
    if (capEl)   { capEl.textContent = fmtSec(capacity);     capEl.dataset.seconds = String(capacity); }
    if (availEl) { availEl.textContent = fmtSec(availability); availEl.dataset.seconds = String(availability); }

    var tkApprovedCard = document.getElementById("capacity-tk-approved-card");
    var idleCard = document.getElementById("capacity-idle-card");
    var tkApprovedEl = document.getElementById("capacity-tk-approved-value");
    var idleEl = document.getElementById("capacity-idle-value");
    var tkApprovedMeta = document.getElementById("capacity-tk-approved-meta");
    var idleMeta = document.getElementById("capacity-idle-meta");
    var analysisMonthEl = document.getElementById("tk-analysis-month-select");
    var selectedMonth = analysisMonthEl ? analysisMonthEl.value : month;
    var tState = monthToggleState();
    if (!tState.anyToggle) {
      if (tkApprovedCard) tkApprovedCard.hidden = true;
      if (idleCard) idleCard.hidden = true;
      return;
    }
    var se = scopedEpics();
    var matched = se.filter(function (e) {
      var s = String(e.start_date || "").slice(0,7);
      var d = String(e.due_date || "").slice(0,7);
      return (tState.started && s === selectedMonth)
        || (tState.delivered && d === selectedMonth)
        || (tState.through && s && d && s <= selectedMonth && selectedMonth <= d);
    });
    var tkApprovedSec = matched.reduce(function (n, e) { return n + Number(e.tk_approved_seconds || 0); }, 0);
    var idleSec = availability - tkApprovedSec;
    if (tkApprovedEl) {
      tkApprovedEl.textContent = fmtSec(tkApprovedSec);
      tkApprovedEl.dataset.seconds = String(tkApprovedSec);
    }
    if (idleEl) {
      idleEl.textContent = fmtSec(idleSec);
      idleEl.dataset.seconds = String(idleSec);
    }
    if (tkApprovedMeta) tkApprovedMeta.textContent = "TK approved for " + monthName(selectedMonth) + " with active month toggles";
    if (idleMeta) {
      var idleDesc = idleSec < 0 ? "TK approved exceeds total availability for selected month" : "Availability minus TK approved for selected month";
      idleMeta.textContent = idleDesc;
    }
    if (tkApprovedCard) tkApprovedCard.hidden = false;
    if (idleCard) idleCard.hidden = false;
  }

  /* ── Month analysis ───────────────────────────────────────────────── */
  function monthName(ym) {
    var parts = ym.split("-");
    var names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return (names[Number(parts[1]) - 1] || "") + " " + parts[0];
  }

  function parseMonthKey(iso) {
    if (!iso) return "";
    var m = /^(\\d{4})-(\\d{2})/.exec(String(iso));
    return m ? m[1] + "-" + m[2] : "";
  }
  function monthKeyRange(startKey, endKey) {
    if (!startKey && !endKey) return [];
    if (!startKey) return endKey ? [endKey] : [];
    if (!endKey) return [startKey];
    var sy = parseInt(startKey.split("-")[0], 10);
    var sm = parseInt(startKey.split("-")[1], 10);
    var ey = parseInt(endKey.split("-")[0], 10);
    var em = parseInt(endKey.split("-")[1], 10);
    if (isNaN(sy) || isNaN(sm) || isNaN(ey) || isNaN(em)) return [];
    var sDate = new Date(sy, sm - 1, 1);
    var eDate = new Date(ey, em - 1, 1);
    if (sDate.getTime() > eDate.getTime()) { var tmp = sDate; sDate = eDate; eDate = tmp; }
    var keys = [];
    var c = new Date(sDate.getFullYear(), sDate.getMonth(), 1);
    var endT = eDate.getTime();
    while (c.getTime() <= endT) {
      var mo = c.getMonth() + 1;
      keys.push(c.getFullYear() + "-" + (mo < 10 ? "0" : "") + mo);
      c.setMonth(c.getMonth() + 1);
    }
    return keys;
  }
  function availableMonthKeysForEpics(epicList, throughEnabled) {
    var set = Object.create(null);
    (epicList || []).forEach(function (e) {
      var sk = parseMonthKey(e.start_date);
      var dk = parseMonthKey(e.due_date);
      if (throughEnabled) {
        monthKeyRange(sk, dk).forEach(function (k) { set[k] = true; });
      } else {
        if (dk) set[dk] = true;
        if (sk) set[sk] = true;
      }
    });
    return Object.keys(set).sort();
  }
  function isCrossMonthRange(item) {
    var a = parseMonthKey(item.start_date);
    var b = parseMonthKey(item.due_date);
    return Boolean(a && b && a !== b);
  }
  function bucketMonthKey(item, allowedKeys) {
    var startKey = parseMonthKey(item.start_date);
    var dueKey = parseMonthKey(item.due_date);
    var resolved = "";
    if (startKey && dueKey && startKey === dueKey) resolved = startKey;
    else if (dueKey) resolved = dueKey;
    else resolved = startKey;
    if (!resolved) return "";
    return allowedKeys.indexOf(resolved) >= 0 ? resolved : "";
  }

  function buildMonthStoryAllocations(scopeEpics, allowedKeys) {
    var totalsByKey = {};
    var i, k;
    for (i = 0; i < allowedKeys.length; i++) { totalsByKey[allowedKeys[i]] = 0; }
    var includedEpics = [];
    var excludedEpics = [];

    (scopeEpics || []).forEach(function (epic) {
      var epicTotals = {};
      for (i = 0; i < allowedKeys.length; i++) { epicTotals[allowedKeys[i]] = 0; }
      var reasons = [];
      var contributed = false;
      var stories = Array.isArray(epic.stories) ? epic.stories : [];
      if (!stories.length) {
        reasons.push("No stories available for month estimate analysis.");
      }
      stories.forEach(function (story) {
        var storyKey = story.story_key || story.title || "Unknown story";
        var storyEstimate = Number(story.estimate_seconds) || 0;
        if (isCrossMonthRange(story)) {
          var subtasks = Array.isArray(story.subtasks) ? story.subtasks : [];
          var usable = subtasks.filter(function (st) {
            return (Number(st.estimate_seconds) || 0) > 0 && Boolean(bucketMonthKey(st, allowedKeys));
          });
          if (!usable.length) {
            reasons.push("Story " + storyKey + " spans multiple months but has no usable subtask estimates in the chart months.");
            return;
          }
          usable.forEach(function (st) {
            var bk = bucketMonthKey(st, allowedKeys);
            var est = Number(st.estimate_seconds) || 0;
            if (!bk || est <= 0) return;
            epicTotals[bk] += est;
            contributed = true;
          });
          return;
        }
        var bk2 = bucketMonthKey(story, allowedKeys);
        if (!bk2 || storyEstimate <= 0) return;
        epicTotals[bk2] += storyEstimate;
        contributed = true;
      });
      if (reasons.length || !contributed) {
        var detailReasons = reasons.slice();
        if (!contributed && !detailReasons.length) {
          detailReasons.push("No story estimates contributed to the displayed months (check story/subtask dates and estimates).");
        }
        excludedEpics.push(Object.assign({}, epic, { _monthStoryExclusionReasons: detailReasons }));
        return;
      }
      for (k in epicTotals) { if (Object.prototype.hasOwnProperty.call(epicTotals, k)) { totalsByKey[k] = (totalsByKey[k] || 0) + epicTotals[k]; } }
      includedEpics.push(epic);
    });
    return { totalsByKey: totalsByKey, includedEpics: includedEpics, excludedEpics: excludedEpics };
  }

  function sortEpicsForMonthList(arr) {
    return (arr || []).slice().sort(function (a, b) {
      return String(a.jira_id || "").localeCompare(String(b.jira_id || ""));
    });
  }

  function fillMonthAnalysisEpicLists(includedEpics, excludedEpics) {
    var inclList = document.querySelector("[data-month-analysis-included-list]");
    var exclList = document.querySelector("[data-month-analysis-excluded-list]");
    var incStrong = document.querySelector("[data-month-analysis-included]");
    var excStrong = document.querySelector("[data-month-analysis-excluded]");
    var inc = sortEpicsForMonthList(includedEpics);
    var exc = sortEpicsForMonthList(excludedEpics);
    if (incStrong) incStrong.textContent = String(inc.length);
    if (excStrong) excStrong.textContent = String(exc.length);
    function heading(e) {
      var key = escHtml(String(e.jira_id || ""));
      var title = escHtml(String(e.title || ""));
      var ju = String(e.jira_url || "").trim();
      var head = ju
        ? '<a class="tk-month-epic-link" href="' + escHtml(ju) + '" target="_blank" rel="noopener noreferrer">' + key + "</a>"
        : "<strong>" + key + "</strong>";
      return '<div class="tk-month-epic-heading">' + head + '<span class="tk-month-analysis-epic-title">' + title + "</span></div>";
    }
    if (inclList) {
      inclList.innerHTML = !inc.length
        ? '<div class="tk-month-analysis-list-empty">None for the current scope.</div>'
        : inc.map(function (e) { return '<div class="tk-month-analysis-epic-row">' + heading(e) + "</div>"; }).join("");
    }
    if (exclList) {
      exclList.innerHTML = !exc.length
        ? '<div class="tk-month-analysis-list-empty">None for the current scope.</div>'
        : exc.map(function (e) {
            var reasons = e._monthStoryExclusionReasons;
            var rHtml = "";
            if (reasons && reasons.length) {
              rHtml = '<div class="tk-month-analysis-exclude-reason">' + reasons.map(function (r) { return escHtml(String(r)); }).join("<br>") + "</div>";
            }
            var rem = formatPlannerRemarkText(e.epics_planner_remarks);
            var pHtml = "";
            if (rem) {
              pHtml = '<div class="tk-month-analysis-planner-remark"><span class="lbl">Epics Planner remarks:</span> ' + escHtml(String(rem)).replace(/\\n/g, "<br>") + "</div>";
            }
            return '<div class="tk-month-analysis-epic-row">' + heading(e) + rHtml + pHtml + "</div>";
          }).join("");
    }
  }

  function setMonthAnalysisScopeReconcile(jiraOnly, chartEpicCount) {
    var el = document.querySelector("[data-month-analysis-scope-reconcile]");
    if (!el) return;
    var totalScoped = scopedEpics().length;
    if (jiraOnly && chartEpicCount < totalScoped) {
      var skipped = totalScoped - chartEpicCount;
      el.hidden = false;
      el.textContent = "Included + excluded (" + chartEpicCount + " epics) counts only Jira-populated rows. Total # of RMI Epics (" + totalScoped + ") is the full product/search scope—" + skipped + ' epic(s) appear only there until you turn off "Only Jira Populated Epics".';
    } else {
      el.hidden = true;
    }
  }

  function renderMonthAnalysis() {
    var analysisMonthEl = document.getElementById("tk-analysis-month-select");
    var monthEl = analysisMonthEl || document.getElementById("tk-month-select");
    if (!monthEl) return;
    var month = monthEl.value;
    var tState = monthToggleState();
    var started = tState.started;
    var delivered = tState.delivered;
    var through = tState.through;
    var jiraOnly = document.getElementById("tk-jira-only-enabled") && document.getElementById("tk-jira-only-enabled").checked;
    var anyToggle = tState.anyToggle;

    if (analysisMonthEl) analysisMonthEl.disabled = !anyToggle;
    var panel = document.getElementById("tk-month-analysis");
    var chartBars = document.querySelector("#tk-month-analysis .tk-month-chart-bars");
    if (!panel || !chartBars) return;

    var statusEl = document.querySelector("[data-tk-month-status]");
    if (statusEl) {
      var statusParts = [];
      statusParts.push(activeProduct === "all" ? "All Products" : activeProduct);
      if (jiraOnly) statusParts.push("Jira populated only");
      statusParts.push(anyToggle ? monthName(month) : "all available months");
      statusEl.textContent = statusParts.join(" \u00b7 ");
    }

    var se = scopedEpics();
    if (jiraOnly) se = se.filter(function (e) { return e.jira_populated; });

    /* All toggles off: one bar per available month, story/subtask month buckets (Epics Planner) */
    if (!anyToggle) {
      if (chartBars) {
        chartBars.classList.add("tk-month-chart-bars--all");
      }
      var allowedKeys = availableMonthKeysForEpics(se, false);
      if (!allowedKeys.length) {
        panel.hidden = true;
        var recEarly = document.querySelector("[data-month-analysis-scope-reconcile]");
        if (recEarly) recEarly.hidden = true;
        return;
      }
      var alloc = buildMonthStoryAllocations(se, allowedKeys);
      var maxSec = 0;
      var peakKey = allowedKeys[0];
      for (var ai = 0; ai < allowedKeys.length; ai++) {
        var k0 = allowedKeys[ai];
        var sec0 = Number(alloc.totalsByKey[k0] || 0);
        if (sec0 > maxSec) { maxSec = sec0; peakKey = k0; }
      }
      if (maxSec < 1e-6) maxSec = 1;

      var h = "";
      for (var bi = 0; bi < allowedKeys.length; bi++) {
        var mk = allowedKeys[bi];
        var sec = Number(alloc.totalsByKey[mk] || 0);
        var feat = mk === peakKey && sec > 0 ? " featured" : "";
        var pct = Math.max(6, (sec / maxSec) * 100);
        h += '<section class="tk-month-chart-bar-card' + feat + '" data-month-analysis-slot="all-' + bi + '">'
          + '<div class="tk-month-chart-value duration-value" data-month-analysis-chart-value data-seconds="' + sec + '">' + fmtSec(sec) + "</div>"
          + '<div class="tk-month-chart-bar-track"><div class="tk-month-chart-bar-fill" data-month-analysis-bar style="height:' + pct + '%"></div></div>'
          + '<div class="tk-month-chart-label" data-month-analysis-label>' + monthName(mk) + "</div></section>";
      }
      chartBars.innerHTML = h;

      fillMonthAnalysisEpicLists(alloc.includedEpics, alloc.excludedEpics);
      setMonthAnalysisScopeReconcile(jiraOnly, se.length);
      panel.hidden = false;
      return;
    }

    if (chartBars) {
      chartBars.classList.remove("tk-month-chart-bars--all");
    }

    function matchMonth(e) {
      var s = String(e.start_date || "").slice(0,7);
      var d = String(e.due_date || "").slice(0,7);
      return (started && s === month) || (delivered && d === month) || (through && s && d && s <= month && month <= d);
    }
    var matched = se.filter(matchMonth);
    var storySec   = matched.reduce(function (n, e) { return n + Number(e.story_estimate_seconds || 0); }, 0);

    var months = monthEl.options;
    var idx = -1;
    for (var j = 0; j < months.length; j++) { if (months[j].value === month) { idx = j; break; } }
    var prevMonth = idx > 0 ? months[idx - 1].value : "";
    var nextMonth = idx < months.length - 1 ? months[idx + 1].value : "";

    function monthTotal(m) {
      if (!m) return 0;
      return se.filter(function (e) {
        var s = String(e.start_date || "").slice(0,7);
        var d = String(e.due_date || "").slice(0,7);
        return (started && s === m) || (delivered && d === m) || (through && s && d && s <= m && m <= d);
      }).reduce(function (n, e) { return n + Number(e.story_estimate_seconds || 0); }, 0);
    }
    var vals = [monthTotal(prevMonth), storySec, monthTotal(nextMonth)];
    var maxVal = Math.max.apply(null, vals) || 1;

    var labels = [prevMonth ? monthName(prevMonth) : "Previous", monthName(month), nextMonth ? monthName(nextMonth) : "Next"];
    var htmlScoped = "";
    for (var si = 0; si < 3; si++) {
      var pvi = (vals[si] / maxVal) * 100;
      var fcls = si === 1 ? " featured" : "";
      htmlScoped += '<section class="tk-month-chart-bar-card' + fcls + '" data-month-analysis-slot="' + ["previous","selected","next"][si] + '">'
        + '<div class="tk-month-chart-value duration-value" data-month-analysis-chart-value data-seconds="' + vals[si] + '">' + fmtSec(vals[si]) + "</div>"
        + '<div class="tk-month-chart-bar-track"><div class="tk-month-chart-bar-fill" data-month-analysis-bar style="height:' + Math.max(6, pvi) + '%"></div></div>'
        + '<div class="tk-month-chart-label" data-month-analysis-label>' + labels[si] + "</div></section>";
    }
    chartBars.innerHTML = htmlScoped;

    var matchedSet = new Set(matched);
    var excludedScoped = se.filter(function (e) { return !matchedSet.has(e); }).map(function (e) {
      return Object.assign({}, e, {
        _monthStoryExclusionReasons: [
          "Epic does not match the active month filters (started in / delivered in / any work done through).",
        ],
      });
    });
    fillMonthAnalysisEpicLists(matched, excludedScoped);
    setMonthAnalysisScopeReconcile(jiraOnly, se.length);
    panel.hidden = false;
  }

  /* ── Table filtering ──────────────────────────────────────────────── */
  function applyFilters() {
    var q = searchQuery();
    var shown = 0;
    document.querySelectorAll(".epic-row").forEach(function (row) {
      var okProduct = activeProduct === "all" || row.dataset.product === activeProduct;
      var okSearch = !q || row.dataset.search.indexOf(q) >= 0;
      var visible = okProduct && okSearch;
      row.hidden = !visible;
      if (!visible) {
        document.querySelectorAll(".child-of-" + CSS.escape(row.dataset.epicId)).forEach(function (c) { c.hidden = true; });
        var btn = row.querySelector(".row-toggle");
        if (btn) btn.textContent = "+";
      }
      if (visible) shown++;
    });

    var status = document.querySelector(".search-status");
    if (status) status.textContent = q || activeProduct !== "all" ? shown + " of " + epics.length + " epics" : "";

    /* product toolbar */
    document.querySelectorAll(".product-toggle").forEach(function (b) {
      b.classList.toggle("active", b.dataset.product === activeProduct);
    });

    renderMetrics();
    renderCapacity();
    renderMonthAnalysis();
    renderGantt();
  }

  /* ── Row expand/collapse ──────────────────────────────────────────── */
  document.addEventListener("click", function (e) {
    var btn = e.target.closest(".row-toggle");
    if (!btn) return;
    var epicId = btn.dataset.epic;
    var storyId = btn.dataset.story;
    var worklogPanelId = btn.dataset.worklogPanel;
    if (epicId) {
      var rows = document.querySelectorAll(".child-of-" + CSS.escape(epicId));
      var show = Array.from(rows).some(function (r) { return r.hidden; });
      rows.forEach(function (r) {
        if (r.classList.contains("story-row")) { r.hidden = !show; }
        else { r.hidden = true; } /* hide subtasks when toggling epic */
      });
      /* reset story toggles */
      rows.forEach(function (r) { var b = r.querySelector(".row-toggle"); if (b) b.textContent = "+"; });
      btn.textContent = show ? "\u2212" : "+";
    } else if (storyId) {
      var subs = document.querySelectorAll(".child-of-" + CSS.escape(storyId));
      var showSub = Array.from(subs).some(function (r) {
        return !r.classList.contains("subtask-worklog-panel") && r.hidden;
      });
      subs.forEach(function (r) {
        if (r.classList.contains("subtask-worklog-panel")) {
          if (!showSub) {
            r.hidden = true;
            var pid = r.getAttribute("data-worklog-panel");
            if (pid) {
              var tbtn = document.querySelector('.row-toggle-subtask[data-worklog-panel="' + CSS.escape(pid) + '"]');
              if (tbtn) {
                tbtn.textContent = "+";
                tbtn.setAttribute("aria-expanded", "false");
              }
            }
          }
          return;
        }
        r.hidden = !showSub;
      });
      btn.textContent = showSub ? "\u2212" : "+";
    } else if (worklogPanelId) {
      var panel = document.querySelector('.subtask-worklog-panel[data-worklog-panel="' + CSS.escape(worklogPanelId) + '"]');
      if (!panel) return;
      var opening = panel.hidden;
      panel.hidden = !opening;
      btn.textContent = opening ? "\u2212" : "+";
      btn.setAttribute("aria-expanded", opening ? "true" : "false");
    }
  });

  /* ── Drawer ───────────────────────────────────────────────────────── */
  function openDrawer(title, items) {
    var overlay = document.getElementById("detail-drawer");
    if (!overlay) return;
    var body = document.getElementById("drawer-body");
    body.innerHTML = '<div class="drawer-epics-list">' + items.map(function (e) {
      return '<div class="drawer-epic-card"><strong>' + (e.jira_id || "") + '</strong> ' + (e.title || "") + '<div class="meta">' + fmtSec(e.value) + '</div></div>';
    }).join("") + '</div>';
    document.getElementById("drawer-title").textContent = title;
    overlay.classList.add("open");
  }

  function escHtml(s) {
    return String(s === undefined || s === null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function formatPlannerRemarkText(value) {
    var raw = String(value === undefined || value === null ? "" : value).trim();
    if (!raw) return "";
    var decoded = raw;
    for (var i = 0; i < 2; i += 1) {
      var ta = document.createElement("textarea");
      ta.innerHTML = decoded;
      decoded = ta.value;
    }
    decoded = decoded
      .replace(/<br\\s*\\/?\\s*>/gi, "\\n")
      .replace(/<li\\b[^>]*>/gi, "- ")
      .replace(/<\\/(p|div|li|tr|h[1-6])>/gi, "\\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\\u00a0/g, " ");
    return decoded
      .split(/\\r?\\n/)
      .map(function (line) { return line.trim(); })
      .filter(Boolean)
      .join("\\n")
      .trim();
  }
  function normEmp(n) { return String(n || "").trim().toLowerCase(); }
  function nameInEmpList(name, arr) {
    var k = normEmp(name);
    return (arr || []).some(function (x) { return normEmp(x) === k; });
  }
  function openHtmlDrawer(title, html) {
    var overlay = document.getElementById("detail-drawer");
    if (!overlay) return;
    document.getElementById("drawer-title").textContent = title;
    document.getElementById("drawer-body").innerHTML = html;
    overlay.classList.add("open");
  }
  function buildEmployeeBreakdownHtml() {
    var month = (document.getElementById("tk-month-select") || {}).value || "";
    var prof = findCapacityProfileForMonth(month);
    var global = isGlobalCapacityScope();
    var considered = scopedCapacityAssignees();
    var consSet = {};
    considered.forEach(function (n) { consSet[normEmp(n)] = true; });
    var cs = capacitySource || {};
    var universe = cs.capacity_universe_assignees || [];
    var orgList = cs.org_team_assignees || [];
    var rltList = cs.rlt_employees || [];
    var atMap = cs.assignee_to_teams || {};

    function rowTable(rows) {
      if (!rows || !rows.length) return '<p class="drawer-empty">None.</p>';
      var h = '<table class="drawer-emp-table"><thead><tr><th class="drawer-emp-idx">#</th><th>Name</th><th>Details</th></tr></thead><tbody>';
      rows.forEach(function (r, i) {
        h += '<tr><td class="drawer-emp-idx">' + (i + 1) + '</td><td>' + escHtml(r.name) + '</td><td>' + escHtml(r.note) + '</td></tr>';
      });
      h += '</tbody></table>';
      return h;
    }

    var parts = [];
    parts.push('<p class="drawer-prose">This follows the same rules as the capacity calculator: month, team selection, product filter, and capacity profile. Click the overlay or Close to dismiss.</p>');

    if (global && prof) {
      var nFte = Math.max(0, Math.round(Number(prof.employee_count) || 0));
      parts.push('<div class="drawer-employee-section"><h3>Headcount source (policy FTEs)</h3>');
      parts.push(
        '<p class="drawer-prose">The value shown in the report is <strong>' + escHtml(String(nFte)) +
        '</strong> FTEs from the capacity profile covering <strong>' + escHtml(prof.from_date || "") +
        '</strong> to <strong>' + escHtml(prof.to_date || "") +
        '</strong> (overlaps month <strong>' + escHtml(month) +
        '</strong>). This is a single org-wide number from the database, not a sum of the named people in the next section.</p></div>'
      );
      var incRoster = orgList.map(function (name) {
        var tms = atMap[name] || [];
        return { name: name, note: tms.length ? "Performance team(s): " + tms.join(", ") : "Performance team member" };
      });
      parts.push('<div class="drawer-employee-section"><h3>Considered (reference — performance teams)</h3>');
      parts.push('<p class="drawer-prose">People mapped to at least one team. When you filter by team, headcount and leave use this roster. The policy FTE total above is separate.</p>');
      parts.push(rowTable(incRoster));
      parts.push('</div>');

      var gap = [];
      rltList.forEach(function (name) {
        if (!nameInEmpList(name, orgList)) {
          gap.push({ name: name, note: "In RLT Jira assignee summary but not on any performance team in the database" });
        }
      });
      orgList.forEach(function (name) {
        if (!nameInEmpList(name, rltList)) {
          gap.push({ name: name, note: "On a performance team but not in the RLT Jira assignee summary for this import" });
        }
      });
      if (gap.length) {
        parts.push('<div class="drawer-employee-section"><h3>Data alignment (not automatic exclusions from FTE)</h3>');
        parts.push('<p class="drawer-prose">These people appear on one list but not the other. This does not remove anyone from the FTE total by name.</p>');
        parts.push(rowTable(gap));
        parts.push('</div>');
      }
      return parts.join("");
    }

    if (!isAllMembersSelected()) {
      var st = selectedTeamNames();
      var inc = considered.map(function (name) {
        var onTeam = (atMap[name] || []).filter(function (x) { return st.indexOf(x) >= 0; });
        return { name: name, note: onTeam.length ? "Selected via team(s): " + onTeam.join(", ") : "Selected in team/member filter" };
      });
      parts.push('<div class="drawer-employee-section"><h3>Considered for this headcount</h3>');
      parts.push('<p class="drawer-prose">Selected members'
        + (st.length ? ' across team(s): ' + escHtml(st.join(", ")) : '.')
        + '</p>');
      parts.push(rowTable(inc));
      parts.push('</div>');

      var ex = [];
      universe.forEach(function (name) {
        if (!consSet[normEmp(name)]) {
          ex.push({ name: name, note: "Not selected in the team/member filter" });
        }
      });
      parts.push('<div class="drawer-employee-section"><h3>Not considered</h3>');
      parts.push(rowTable(ex));
      parts.push('</div>');
      return parts.join("");
    }

    if (activeProduct !== "all") {
      var pl = activeProduct;
      var inc2 = considered.map(function (name) {
        return { name: name, note: "Story or subtask assignee under an RMI epic in this product" };
      });
      parts.push('<div class="drawer-employee-section"><h3>Considered for this headcount</h3>');
      parts.push('<p class="drawer-prose">Product: <strong>' + escHtml(pl) + '</strong>. Only people with assigned work in the loaded RMI tree for this product.</p>');
      parts.push(rowTable(inc2));
      parts.push('</div>');

      var ex2 = [];
      universe.forEach(function (name) {
        if (!consSet[normEmp(name)]) {
          ex2.push({ name: name, note: "No story or subtask in this product in the RMI set with this assignee (or work is unassigned only)" });
        }
      });
      parts.push('<div class="drawer-employee-section"><h3>Not considered</h3>');
      parts.push(rowTable(ex2));
      parts.push('</div>');
      return parts.join("");
    }

    if (global && !prof) {
      var inc3 = considered.map(function (name) {
        return { name: name, note: "In the RLT assignee list used when no profile covers the month" };
      });
      parts.push('<div class="drawer-employee-section"><h3>Considered for this headcount</h3>');
      parts.push('<p class="drawer-prose">No capacity profile date range fully overlaps the selected month. The headcount uses RLT Jira assignees from the canonical run.</p>');
      parts.push(rowTable(inc3));
      parts.push('</div>');

      var ex3 = [];
      universe.forEach(function (name) {
        if (!consSet[normEmp(name)]) {
          ex3.push({ name: name, note: "Not in that RLT assignee list" });
        }
      });
      parts.push('<div class="drawer-employee-section"><h3>Not considered</h3>');
      parts.push(rowTable(ex3));
      parts.push('</div>');
      return parts.join("");
    }

    var inc4 = considered.map(function (name) {
      return { name: name, note: "Included in the active headcount list" };
    });
    parts.push('<div class="drawer-employee-section"><h3>Considered for this headcount</h3>');
    parts.push(rowTable(inc4));
    parts.push('</div>');

    var ex4 = [];
    universe.forEach(function (name) {
      if (!consSet[normEmp(name)]) {
        ex4.push({ name: name, note: "Outside the assignee set used for the capacity row in this view" });
      }
    });
    parts.push('<div class="drawer-employee-section"><h3>Not considered</h3>');
    parts.push(rowTable(ex4));
    parts.push('</div>');
    return parts.join("");
  }
  function showEmployeeBreakdownDrawer() {
    openHtmlDrawer("No. of Employees \u2014 breakdown", buildEmployeeBreakdownHtml());
  }

  document.getElementById("drawer-close").addEventListener("click", function () {
    document.getElementById("detail-drawer").classList.remove("open");
  });
  document.getElementById("detail-drawer").addEventListener("click", function (e) {
    if (e.target === document.getElementById("detail-drawer")) {
      document.getElementById("detail-drawer").classList.remove("open");
    }
  });

  (function initEmployeeHeadcountDrawer() {
    var el = document.getElementById("capacity-employees-val");
    if (!el) return;
    el.addEventListener("click", function (e) {
      e.stopPropagation();
      e.preventDefault();
      showEmployeeBreakdownDrawer();
    });
    el.addEventListener("keydown", function (e) {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        e.stopPropagation();
        showEmployeeBreakdownDrawer();
      }
    });
  })();

  document.getElementById("metric-grid").addEventListener("click", function (e) {
    var card = e.target.closest(".metric-card-clickable");
    if (!card) return;
    var key = card.dataset.metricKey;
    var label = card.querySelector(".metric-label").textContent;
    var items = scopedEpics().map(function (ep) {
      return {jira_id: ep.jira_id, title: ep.title, value: Number(ep[key] || 0)};
    }).filter(function (x) { return x.value > 0; });
    items.sort(function (a, b) { return b.value - a.value; });
    openDrawer(label, items);
  });

  /* ── Gantt ────────────────────────────────────────────────────────── */
  function renderGantt() {
    var container = document.getElementById("gantt-view");
    if (!container) return;
    var dated = scopedEpics().filter(function (e) { return e.start_date && e.due_date; });
    if (!dated.length) { container.innerHTML = '<div class="empty-state">No dated epics in scope.</div>'; return; }

    var products = [];
    var seen = {};
    dated.forEach(function (e) { var p = e.product || "Unassigned"; if (!seen[p]) { seen[p] = true; products.push(p); } });
    products.sort();

    var allStarts = dated.map(function (e) { return new Date(e.start_date + "T00:00:00").getTime(); });
    var allEnds   = dated.map(function (e) { return new Date(e.due_date   + "T00:00:00").getTime(); });
    var globalMin = Math.min.apply(null, allStarts);
    var globalMax = Math.max.apply(null, allEnds);
    var span = Math.max(1, globalMax - globalMin);

    var rowH = 32, padY = 8, headerH = 30;
    var html = '<div class="gantt-product-grid">';

    products.forEach(function (p, pi) {
      var pEpics = dated.filter(function (e) { return (e.product || "Unassigned") === p; });
      var svgH = headerH + pEpics.length * (rowH + padY) + padY;
      var color = PRODUCT_COLORS[pi % PRODUCT_COLORS.length];
      html += '<div class="gantt-product-section"><h3 style="margin:0 0 8px;color:' + color + '">' + p + ' (' + pEpics.length + ')</h3>'
            + '<svg width="100%" height="' + svgH + '" style="display:block">';
      /* month grid lines */
      var d = new Date(globalMin);
      d.setDate(1);
      while (d.getTime() <= globalMax) {
        var x = ((d.getTime() - globalMin) / span) * 100;
        var label = d.toLocaleString("default", {month:"short",year:"2-digit"});
        html += '<line x1="' + x + '%" y1="0" x2="' + x + '%" y2="' + svgH + '" stroke="#d0dbe6" stroke-dasharray="4,4"/>';
        html += '<text x="' + x + '%" y="16" class="gantt-meta" dx="4">' + label + '</text>';
        d.setMonth(d.getMonth() + 1);
      }
      pEpics.forEach(function (e, ei) {
        var s = new Date(e.start_date + "T00:00:00").getTime();
        var en = new Date(e.due_date + "T00:00:00").getTime();
        var x = ((s - globalMin) / span) * 100;
        var w = Math.max(1, ((en - s) / span) * 100);
        var y = headerH + ei * (rowH + padY) + padY;
        html += '<rect x="' + x + '%" y="' + y + '" width="' + w + '%" height="' + rowH + '" rx="6" fill="' + color + '" opacity=".82"/>';
        html += '<text x="' + x + '%" y="' + (y + 20) + '" class="gantt-label" dx="8" fill="#fff" style="font-size:11px">' + e.jira_id + '</text>';
      });
      html += '</svg></div>';
    });
    html += '</div>';
    container.innerHTML = html;
  }

  /* ── View toggle ──────────────────────────────────────────────────── */
  document.querySelectorAll(".view-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var target = btn.dataset.view;
      document.getElementById("table-view-section").hidden = (target !== "table");
      document.getElementById("gantt-view-section").hidden  = (target !== "gantt");
      document.querySelectorAll(".view-toggle").forEach(function (b) { b.classList.toggle("active", b === btn); });
      if (target === "gantt") renderGantt();
    });
  });

  /* ── Unit toggle ──────────────────────────────────────────────────── */
  document.querySelectorAll(".unit-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      unit = btn.dataset.unit;
      document.querySelectorAll(".unit-toggle").forEach(function (b) { b.classList.toggle("active", b === btn); });
      renderMetrics();
      updateAllDurations();
      if (typeof renderRmiProductCards === "function") renderRmiProductCards();
      if (typeof renderRmiScheduleTable === "function") renderRmiScheduleTable();
    });
  });

  /* ── Product toolbar ──────────────────────────────────────────────── */
  document.querySelectorAll(".product-toggle").forEach(function (btn) {
    btn.addEventListener("click", function () {
      activeProduct = btn.dataset.product;
      applyFilters();
    });
  });

  /* ── Search ───────────────────────────────────────────────────────── */
  var searchEl = document.getElementById("epic-search");
  if (searchEl) {
    searchEl.addEventListener("input", applyFilters);
  }
  var clearEl = document.getElementById("search-clear");
  if (clearEl) {
    clearEl.addEventListener("click", function () {
      if (searchEl) { searchEl.value = ""; applyFilters(); }
    });
  }

  /* ── Month select (sync capacity + analysis selectors) ────────────── */
  var monthSel = document.getElementById("tk-month-select");
  var analysisMonthSel = document.getElementById("tk-analysis-month-select");
  function syncMonthSelectors(source) {
    var v = source.value;
    if (monthSel && monthSel !== source) monthSel.value = v;
    if (analysisMonthSel && analysisMonthSel !== source) analysisMonthSel.value = v;
    renderCapacity(); renderMonthAnalysis();
  }
  if (monthSel) monthSel.addEventListener("change", function () { syncMonthSelectors(monthSel); });
  if (analysisMonthSel) analysisMonthSel.addEventListener("change", function () { syncMonthSelectors(analysisMonthSel); });
  ["tk-start-month-enabled","tk-month-enabled","tk-through-month-enabled","tk-jira-only-enabled"].forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.addEventListener("change", function () { renderCapacity(); renderMonthAnalysis(); });
  });
  var diagnosticsToggle = document.getElementById("diagnostics-toggle-enabled");
  if (diagnosticsToggle) diagnosticsToggle.addEventListener("change", function () { renderMetrics(); });

  /* ── RMI Estimation & Scheduling (IPP reference layout) ─────────── */
  var rmiScheduleRecords = DATA.rmi_schedule_records || [];
  var rmiScheduleYears = DATA.rmi_schedule_years || [];
  var rmiScheduleBody = document.getElementById("rmi-schedule-body");
  var rmiScheduleFoot = document.getElementById("rmi-schedule-foot");
  var rmiScheduleYearSelect = document.getElementById("rmi-schedule-year");
  var rmiProductCardsContainer = document.getElementById("rmi-sched-product-cards");
  var rmiJiraOnlyToggle = document.getElementById("rmi-jira-only-toggle");
  var rmiSelectedYear = new Date().getFullYear();
  var rmiActiveProduct = "all";
  var rmiJiraOnly = false;

  function rmiScheduleProductColor(product) {
    return PRODUCT_ACCENTS[product] || "#475569";
  }

  function formatMetricDurationCompact(seconds, u) {
    var value = u === "days" ? Number(seconds) / SEC_PER_DAY : Number(seconds) / SEC_PER_HOUR;
    var suffix = u === "days" ? " d" : " h";
    return Math.round(value).toLocaleString() + suffix;
  }

  function rmiFilteredRecords() {
    var recs = rmiScheduleRecords;
    if (rmiJiraOnly) recs = recs.filter(function (e) { return e.jira_populated; });
    if (rmiActiveProduct !== "all") recs = recs.filter(function (e) { return e.product === rmiActiveProduct; });
    return recs;
  }

  function rmiScheduleInit() {
    if (!rmiScheduleYearSelect) return;
    var years = rmiScheduleYears.slice();
    if (!years.length) years.push(new Date().getFullYear());
    rmiSelectedYear = years[years.length - 1];
    rmiScheduleYearSelect.innerHTML = "";
    years.forEach(function (year) {
      var opt = document.createElement("option");
      opt.value = String(year);
      opt.textContent = String(year);
      if (year === rmiSelectedYear) opt.selected = true;
      rmiScheduleYearSelect.appendChild(opt);
    });
    rmiScheduleYearSelect.addEventListener("change", function () {
      rmiSelectedYear = Number(rmiScheduleYearSelect.value);
      renderRmiScheduleTable();
    });
    if (rmiJiraOnlyToggle) {
      rmiJiraOnly = rmiJiraOnlyToggle.checked;
      rmiJiraOnlyToggle.addEventListener("change", function () {
        rmiJiraOnly = rmiJiraOnlyToggle.checked;
        renderRmiProductCards();
        renderRmiScheduleTable();
      });
    }
    renderRmiProductCards();
    renderRmiScheduleTable();
    var headerGroupRow = document.querySelector(".rmi-sched-header-groups");
    var schedTable = document.querySelector(".rmi-schedule-table");
    if (headerGroupRow && schedTable) {
      var h = headerGroupRow.getBoundingClientRect().height;
      schedTable.style.setProperty("--rmi-sched-row1-h", h + "px");
    }
  }

  function renderRmiProductCards() {
    if (!rmiProductCardsContainer) return;
    var base = rmiJiraOnly ? rmiScheduleRecords.filter(function (e) { return e.jira_populated; }) : rmiScheduleRecords;
    var tkByProduct = {};
    var countByProduct = {};
    var totalTk = 0;
    base.forEach(function (epic) {
      var tk = (Number(epic.tk_approved_days) || 0) * SEC_PER_DAY;
      var p = epic.product || "Unassigned";
      tkByProduct[p] = (tkByProduct[p] || 0) + tk;
      countByProduct[p] = (countByProduct[p] || 0) + 1;
      totalTk += tk;
    });
    var products = Object.keys(tkByProduct).sort();
    var totalCount = base.length;
    var allColor = "#102033";
    var allHours = formatMetricDurationCompact(totalTk, unit);
    var html = '<section class="rmi-sched-pcard active" data-rmi-pcard="all" style="--product-accent:' + allColor + '" role="button" tabindex="0" aria-pressed="true">';
    html += '<div class="rmi-sched-pcard-label">All Products</div>';
    html += '<div class="rmi-sched-pcard-value duration-value" data-seconds="' + totalTk + '">' + escHtml(allHours) + "</div>";
    html += '<div class="rmi-sched-pcard-meta">' + totalCount.toLocaleString() + " RMIs/Epics \u2022 Total TK Approved</div></section>";
    products.forEach(function (product) {
      var color = rmiScheduleProductColor(product);
      var seconds = tkByProduct[product] || 0;
      var count = countByProduct[product] || 0;
      var hText = formatMetricDurationCompact(seconds, unit);
      html += '<section class="rmi-sched-pcard" data-rmi-pcard="' + escHtml(product) + '" style="--product-accent:' + color + '" role="button" tabindex="0" aria-pressed="false">';
      html += '<div class="rmi-sched-pcard-label">' + escHtml(product) + "</div>";
      html += '<div class="rmi-sched-pcard-value duration-value" data-seconds="' + seconds + '">' + escHtml(hText) + "</div>";
      html += '<div class="rmi-sched-pcard-meta">' + count.toLocaleString() + " RMIs/Epics \u2022 Total TK Approved</div></section>";
    });
    rmiProductCardsContainer.innerHTML = html;
    rmiProductCardsContainer.querySelectorAll(".rmi-sched-pcard").forEach(function (card) {
      function handleClick() {
        rmiActiveProduct = card.getAttribute("data-rmi-pcard") || "all";
        rmiProductCardsContainer.querySelectorAll(".rmi-sched-pcard").forEach(function (c) {
          var isActive = c.getAttribute("data-rmi-pcard") === rmiActiveProduct;
          c.classList.toggle("active", isActive);
          c.setAttribute("aria-pressed", String(isActive));
        });
        renderRmiScheduleTable();
      }
      card.addEventListener("click", handleClick);
      card.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          handleClick();
        }
      });
    });
  }

  function rmiBucketEpicMonths(epic) {
    var totals = {};
    var stories = Array.isArray(epic.stories) ? epic.stories : [];
    stories.forEach(function (story) {
      var storyEstimate = Number(story.estimate_seconds) || 0;
      if (isCrossMonthRange(story)) {
        var subtasks = Array.isArray(story.subtasks) ? story.subtasks : [];
        subtasks.forEach(function (subtask) {
          var est = Number(subtask.estimate_seconds) || 0;
          if (est <= 0) return;
          var sk = parseMonthKey(subtask.start_date);
          var dk = parseMonthKey(subtask.due_date);
          var mk = "";
          if (sk && dk && sk === dk) mk = sk;
          else if (dk) mk = dk;
          else mk = sk;
          if (!mk) return;
          totals[mk] = (totals[mk] || 0) + est;
        });
        return;
      }
      if (storyEstimate <= 0) return;
      var sk2 = parseMonthKey(story.start_date);
      var dk2 = parseMonthKey(story.due_date);
      var mk2 = "";
      if (sk2 && dk2 && sk2 === dk2) mk2 = sk2;
      else if (dk2) mk2 = dk2;
      else mk2 = sk2;
      if (!mk2) return;
      totals[mk2] = (totals[mk2] || 0) + storyEstimate;
    });
    return totals;
  }

  function rmiFormatValue(seconds) {
    if (!seconds) return "";
    if (unit === "days") return Math.round(seconds / SEC_PER_DAY).toLocaleString();
    return Math.round(seconds / SEC_PER_HOUR).toLocaleString();
  }

  function rmiFormatDays(days) {
    var visible = rmiDisplayDays(days);
    return visible ? visible.toLocaleString() : "";
  }

  function rmiDisplayDays(days) {
    if (!days) return 0;
    if (unit === "days") return Math.round(Number(days));
    return Math.round(Number(days) * 8);
  }

  function rmiStatusDataLower(status) {
    return String(status || "")
      .toLowerCase()
      .replace(/_/g, " ")
      .replace(/-/g, " ")
      .trim();
  }

  function renderRmiScheduleTable() {
    if (!rmiScheduleBody || !rmiScheduleFoot) return;
    var yearStr = String(rmiSelectedYear);
    var monthKeys = [];
    for (var mi = 0; mi < 12; mi++) {
      var mo = mi + 1;
      monthKeys.push(yearStr + "-" + (mo < 10 ? "0" : "") + mo);
    }
    var filtered = rmiFilteredRecords();
    var productOrder = [];
    var byProduct = {};
    filtered.forEach(function (epic) {
      var p = epic.product || "Unassigned";
      if (!byProduct[p]) {
        byProduct[p] = [];
        productOrder.push(p);
      }
      byProduct[p].push(epic);
    });
    productOrder.sort();
    var bodyHtml = "";
    var grandTotalMonths = [0,0,0,0,0,0,0,0,0,0,0,0];
    var grandMl = 0;
    var grandTk = 0;
    productOrder.forEach(function (product) {
      var color = rmiScheduleProductColor(product);
      bodyHtml += '<tr class="rmi-sched-product-group"><td></td><td class="rmi-sched-group-label" style="border-left-color:' + color + '">' + escHtml(product) + "</td>";
      for (var gi = 0; gi < 16; gi++) bodyHtml += "<td></td>";
      bodyHtml += "</tr>";
      var epics = byProduct[product];
      var subtotalMonths = [0,0,0,0,0,0,0,0,0,0,0,0];
      var subtotalMl = 0;
      var subtotalTk = 0;
      var rowNum = 0;
      epics.forEach(function (epic) {
        rowNum++;
        var buckets = rmiBucketEpicMonths(epic);
        var stLabel = epic.status ? epic.status : "\u2014";
        var statusLower = rmiStatusDataLower(epic.status);
        var ju = epic.jira_url || "";
        var jiraLink = ju && ju !== "#" ? ' <a class="rmi-sched-jira-link" href="' + escHtml(ju) + '" target="_blank" rel="noopener" title="Open in Jira">J</a>' : "";
        var cells = "<td>" + rowNum + "</td>";
        cells += '<td class="rmi-sched-cell-rmi">' + escHtml(epic.roadmap_item) + jiraLink + "</td>";
        cells += "<td>" + escHtml(epic.product) + "</td>";
        cells += '<td><span class="rmi-sched-status-pill" data-status-lower="' + escHtml(statusLower) + '">' + escHtml(stLabel) + "</span></td>";
        cells += "<td>" + rmiFormatDays(epic.most_likely_days) + "</td>";
        cells += "<td>" + rmiFormatDays(epic.tk_approved_days) + "</td>";
        subtotalMl += rmiDisplayDays(epic.most_likely_days);
        subtotalTk += rmiDisplayDays(epic.tk_approved_days);
        monthKeys.forEach(function (mk, idx) {
          var val = buckets[mk] || 0;
          subtotalMonths[idx] += val;
          cells += "<td>" + rmiFormatValue(val) + "</td>";
        });
        bodyHtml += '<tr class="rmi-sched-epic-row" data-product="' + escHtml(epic.product) + '">' + cells + "</tr>";
      });
      var subtotalCells = '<td></td><td style="border-left-color:' + color + '">' + escHtml(product) + " Subtotal</td><td></td><td></td>";
      subtotalCells += "<td>" + (subtotalMl ? subtotalMl.toLocaleString() : "") + "</td>";
      subtotalCells += "<td>" + (subtotalTk ? subtotalTk.toLocaleString() : "") + "</td>";
      subtotalMonths.forEach(function (val, idx) {
        grandTotalMonths[idx] += val;
        subtotalCells += "<td>" + rmiFormatValue(val) + "</td>";
      });
      grandMl += subtotalMl;
      grandTk += subtotalTk;
      bodyHtml += '<tr class="rmi-sched-product-subtotal">' + subtotalCells + "</tr>";
    });
    rmiScheduleBody.innerHTML = bodyHtml;
    var totalEpics = filtered.length;
    var footCells = "<td>" + totalEpics + "</td><td>Grand Total</td><td></td><td></td>";
    footCells += "<td>" + (grandMl ? grandMl.toLocaleString() : "") + "</td>";
    footCells += "<td>" + (grandTk ? grandTk.toLocaleString() : "") + "</td>";
    grandTotalMonths.forEach(function (val) {
      footCells += "<td>" + rmiFormatValue(val) + "</td>";
    });
    rmiScheduleFoot.innerHTML = '<tr class="rmi-sched-grand-total">' + footCells + "</tr>";
  }

  /* ── Init ─────────────────────────────────────────────────────────── */
  applyFilters();
  rmiScheduleInit();
})();
"""


def render_html(data: dict[str, Any]) -> str:
    epics = data.get("epics", [])
    products = sorted({_to_text(epic.get("product")) or "Unassigned" for epic in epics})
    months = _available_months(epics)

    PRODUCT_PALETTE = ["#0f766e", "#1d4ed8", "#b45309", "#be123c", "#6d28d9", "#047857", "#4338ca", "#334155"]

    # ── Month option HTML ────────────────────────────────────────────
    MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    def _month_label(ym: str) -> str:
        parts = ym.split("-")
        if len(parts) == 2 and parts[1].isdigit():
            idx = int(parts[1]) - 1
            if 0 <= idx < 12:
                return f"{MONTH_NAMES[idx]} {parts[0]}"
        return ym

    month_options = "".join(
        f'<option value="{escape(m)}"{" selected" if i == len(months) - 1 else ""}>{escape(_month_label(m))}</option>'
        for i, m in enumerate(months)
    )
    if not month_options:
        month_options = '<option value="">No dates</option>'

    rmi_schedule_month_headers = "".join(
        f'<th class="rmi-sched-month" data-month-index="{i + 1}">{escape(MONTH_NAMES[i])}</th>'
        for i in range(12)
    )

    # ── Product toolbar buttons ──────────────────────────────────────
    product_toggles = ['<button class="product-toggle active" data-product="all">All</button>']
    for i, p in enumerate(products):
        product_toggles.append(f'<button class="product-toggle" data-product="{escape(p)}">{escape(p)}</button>')

    # ── Hierarchy table rows ─────────────────────────────────────────
    rows_html_parts: list[str] = []
    for epic in epics:
        eid = _to_text(epic.get("jira_id"))
        stories = epic.get("stories", [])
        has_children = bool(stories)
        toggle = (
            f'<button class="row-toggle" data-epic="{escape(eid)}">+</button>'
            if has_children else '<button class="row-toggle" disabled>&nbsp;</button>'
        )
        search_text = (eid + " " + _to_text(epic.get("title"))).lower()
        rows_html_parts.append(
            f'<tr class="epic-row" data-product="{escape(_to_text(epic.get("product")))}" '
            f'data-epic-id="{escape(eid)}" data-search="{escape(search_text)}">'
            f'<td>{toggle}</td>'
            f'<td><a class="jira-link" href="{escape(_to_text(epic.get("jira_url")))}" target="_blank" rel="noreferrer">{escape(eid)}</a></td>'
            f'<td class="epic-title-text">{escape(_to_text(epic.get("title")))}</td>'
            f'<td>{escape(_to_text(epic.get("product")))}</td>'
            f'<td>{escape(_to_text(epic.get("status")))}</td>'
            f'<td>{escape(_to_text(epic.get("start_date")))}</td>'
            f'<td>{escape(_to_text(epic.get("due_date")))}</td>'
            f'<td class="duration-value" data-seconds="{_to_float(epic.get("tk_approved_seconds"))}">{escape(_duration(epic.get("tk_approved_seconds")))}</td>'
            f'<td class="duration-value" data-seconds="{_to_float(epic.get("story_estimate_seconds"))}">{escape(_duration(epic.get("story_estimate_seconds")))}</td>'
            f'<td class="duration-value" data-seconds="{_to_float(epic.get("subtask_estimate_seconds"))}">{escape(_duration(epic.get("subtask_estimate_seconds")))}</td>'
            f'<td class="duration-value" data-seconds="{_to_float(epic.get("logged_seconds"))}">{escape(_duration(epic.get("logged_seconds")))}</td>'
            f'<td>{escape(str(len(stories)))}</td>'
            "</tr>"
        )
        for story in stories:
            sid = _to_text(story.get("story_key"))
            subtasks = story.get("subtasks", [])
            has_subs = bool(subtasks)
            s_toggle = (
                f'<button class="row-toggle" data-story="{escape(sid)}">+</button>'
                if has_subs else '<button class="row-toggle" disabled>&nbsp;</button>'
            )
            rows_html_parts.append(
                f'<tr class="story-row child-of-{escape(eid)}" hidden>'
                f'<td>{s_toggle}</td>'
                f'<td><a class="jira-link" href="{escape(_to_text(story.get("jira_url")))}" target="_blank" rel="noreferrer">{escape(sid)}</a></td>'
                f'<td class="story-title-text">{escape(_to_text(story.get("title")))}</td>'
                f'<td></td>'
                f'<td>{escape(_to_text(story.get("status")))}</td>'
                f'<td>{escape(_to_text(story.get("start_date")))}</td>'
                f'<td>{escape(_to_text(story.get("due_date")))}</td>'
                f'<td></td>'
                f'<td class="duration-value" data-seconds="{_to_float(story.get("estimate_seconds"))}">{escape(_duration(story.get("estimate_seconds")))}</td>'
                f'<td></td>'
                f'<td class="duration-value" data-seconds="{_to_float(story.get("logged_seconds"))}">{escape(_duration(story.get("logged_seconds")))}</td>'
                f'<td>{escape(str(len(subtasks)))}</td>'
                "</tr>"
            )
            for sub in subtasks:
                sub_key = _to_text(sub.get("issue_key"))
                row_class = "bug-row" if _to_text(sub.get("issue_type")).lower() == "bug" else "subtask-row"
                wl = sub.get("worklogs") if isinstance(sub.get("worklogs"), list) else []
                wl_count = len(wl)
                wl_panel_id = f"{sid}__{sub_key}".replace(" ", "_")
                if wl_count:
                    sub_toggle = (
                        f'<button type="button" class="row-toggle row-toggle-subtask" data-worklog-panel="{escape(wl_panel_id)}" '
                        f'aria-expanded="false" aria-label="Toggle worklogs for {escape(sub_key)}">+</button>'
                    )
                    wl_summary = f"{wl_count} worklog{'s' if wl_count != 1 else ''}"
                else:
                    sub_toggle = '<span class="row-toggle-spacer" aria-hidden="true"></span>'
                    wl_summary = ""
                rows_html_parts.append(
                    f'<tr class="{row_class} child-of-{escape(sid)}" hidden>'
                    f"<td>{sub_toggle}</td>"
                    f'<td><a class="jira-link" href="{escape(_to_text(sub.get("jira_url")))}" target="_blank" rel="noreferrer">{escape(sub_key)}</a></td>'
                    f'<td class="subtask-title-text">{escape(_to_text(sub.get("title")))}</td>'
                    f'<td></td>'
                    f'<td>{escape(_to_text(sub.get("status")))}</td>'
                    f'<td>{escape(_to_text(sub.get("start_date")))}</td>'
                    f'<td>{escape(_to_text(sub.get("due_date")))}</td>'
                    f'<td></td>'
                    f'<td class="duration-value" data-seconds="{_to_float(sub.get("estimate_seconds"))}">{escape(_duration(sub.get("estimate_seconds")))}</td>'
                    f'<td></td>'
                    f'<td class="duration-value" data-seconds="{_to_float(sub.get("logged_seconds"))}">{escape(_duration(sub.get("logged_seconds")))}</td>'
                    f"<td>{escape(wl_summary)}</td>"
                    "</tr>"
                )
                if wl_count:
                    rows_html_parts.append(
                        f'<tr class="subtask-worklog-panel child-of-{escape(sid)}" data-worklog-panel="{escape(wl_panel_id)}" hidden>'
                        '<td colspan="12" class="subtask-worklog-panel-cell">'
                        '<table class="worklog-nested-table">'
                        "<thead><tr>"
                        "<th>Worklog ID</th><th>Author</th><th>Started</th><th>Time Spent</th><th>Hours</th>"
                        "</tr></thead><tbody>"
                        f"{_html_worklog_detail_rows(wl)}"
                        "</tbody></table></td></tr>"
                    )

    table_body = "\n".join(rows_html_parts) if rows_html_parts else '<tr><td colspan="12" class="empty-state">No Epics Planner rows found.</td></tr>'

    # ── Assemble HTML ────────────────────────────────────────────────
    generated_at = escape(_to_text(data.get("generated_at")))
    db_path_text = escape(_to_text(data.get("database_path")))
    canonical_db_text = escape(_to_text(data.get("canonical_database_path")) or _to_text(data.get("database_path")))
    run_id_text = escape(_to_text(data.get("canonical_run_id")) or "none")
    payload_json = json.dumps(data, ensure_ascii=True)

    parts: list[str] = []
    parts.append('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>RMI Jira Gantt</title>')
    parts.append('<link rel="stylesheet" href="shared-nav.css">')
    parts.append("<style>" + _REPORT_CSS + "</style>")
    parts.append("</head>\n<body>\n<div class=\"page\">")

    # Header
    parts.append(f"""
    <header>
      <h1>RMI Jira Gantt</h1>
      <div class="subtext"><strong>Showing TK Epics only.</strong></div>
      <div class="subtext">Generated at {generated_at} from Epics Planner (SQLite) joined with canonical Jira issues/worklogs.</div>
      <div class="subtext" style="margin-top:4px">Epics Planner DB: {db_path_text} | Canonical DB: {canonical_db_text} | Run: {run_id_text}</div>
    </header>""")

    # Metric grid
    parts.append('    <section class="metric-grid" id="metric-grid"></section>')

    # Product summary grid
    parts.append('    <section class="product-summary-grid" id="product-summary-grid"></section>')

    # TK Month filter toolbar
    parts.append(f"""
    <div class="tk-month-toolbar" role="group" aria-label="TK Approved month filter">
      <label class="tk-month-toggle" for="tk-start-month-enabled">
        <input id="tk-start-month-enabled" class="tk-month-toggle-input" type="checkbox">
        <span class="tk-month-toggle-track"><span class="tk-month-toggle-thumb"></span></span>
        <span class="tk-month-toggle-text">For epics started in</span>
      </label>
      <label class="tk-month-toggle" for="tk-month-enabled">
        <input id="tk-month-enabled" class="tk-month-toggle-input" type="checkbox">
        <span class="tk-month-toggle-track"><span class="tk-month-toggle-thumb"></span></span>
        <span class="tk-month-toggle-text">For epics delivered in</span>
      </label>
      <label class="tk-month-toggle" for="tk-through-month-enabled">
        <input id="tk-through-month-enabled" class="tk-month-toggle-input" type="checkbox">
        <span class="tk-month-toggle-track"><span class="tk-month-toggle-thumb"></span></span>
        <span class="tk-month-toggle-text">Any Work Done Through</span>
      </label>
      <select id="tk-analysis-month-select" class="tk-month-select" aria-label="Target analysis month">{month_options}</select>
      <label class="tk-month-toggle" for="tk-jira-only-enabled">
        <input id="tk-jira-only-enabled" class="tk-month-toggle-input" type="checkbox" checked>
        <span class="tk-month-toggle-track"><span class="tk-month-toggle-thumb"></span></span>
        <span class="tk-month-toggle-text">Only Jira Populated Epics</span>
      </label>
      <label class="tk-month-toggle" for="diagnostics-toggle-enabled">
        <input id="diagnostics-toggle-enabled" class="tk-month-toggle-input" type="checkbox">
        <span class="tk-month-toggle-track"><span class="tk-month-toggle-thumb"></span></span>
        <span class="tk-month-toggle-text">Diagnostics</span>
      </label>
      <span class="tk-month-status" data-tk-month-status></span>
    </div>""")

    # Capacity Calculator
    parts.append(f"""
    <div class="capacity-calculator" role="group" aria-label="Capacity calculator">
      <div class="capacity-calc-title">Capacity Calculator</div>
      <div class="capacity-calc-body">
        <div class="capacity-field capacity-team-field">
          <span class="capacity-field-label">Team</span>
          <div class="capacity-ms" id="capacity-team-ms">
            <button type="button" class="capacity-ms-trigger" id="capacity-team-ms-trigger"
              aria-haspopup="listbox" aria-expanded="false" aria-controls="capacity-team-ms-panel" aria-label="Filter by team">
              <span class="capacity-ms-trigger-label" id="capacity-team-ms-label">All Employees</span>
              <span class="capacity-ms-chevron" aria-hidden="true"></span>
            </button>
            <div class="capacity-ms-panel" id="capacity-team-ms-panel" hidden>
              <div class="capacity-ms-search-wrap">
                <input type="search" class="capacity-ms-search" id="capacity-team-ms-search" placeholder="Search teams or members…" autocomplete="off" aria-label="Search teams or members" />
              </div>
              <div class="capacity-ms-list" id="capacity-team-ms-list" role="listbox" aria-multiselectable="true">
                <label class="capacity-ms-item capacity-ms-item-all" id="capacity-team-ms-all-label" data-team-name="__all__">
                  <input type="checkbox" class="capacity-ms-all" id="capacity-team-ms-all" checked />
                  <span class="capacity-ms-item-text" id="capacity-team-ms-all-text">Select all</span>
                </label>
              </div>
              <div class="capacity-ms-actions">
                <button type="button" class="capacity-ms-link" id="capacity-team-ms-clear">Clear selection</button>
              </div>
            </div>
          </div>
        </div>
        <div class="capacity-field">
          <span class="capacity-field-label">No. of Employees</span>
          <div class="capacity-field-value capacity-field-value--clickable" id="capacity-employees-val" role="button" tabindex="0" title="View who is included in this headcount" aria-label="View employee breakdown for headcount">0</div>
        </div>
        <div class="capacity-field">
          <span class="capacity-field-label">Month</span>
          <select id="tk-month-select" class="tk-month-select" aria-label="Target month">{month_options}</select>
        </div>
        <div class="capacity-field">
          <span class="capacity-field-label">Total Leaves (RLT)</span>
          <div class="capacity-field-value" id="capacity-leaves-val">0 h</div>
        </div>
        <div class="capacity-results-grid">
          <section class="metric-card metric-card-capacity capacity-result-card">
            <div class="metric-label">Total Capacity</div>
            <div class="metric-value-wrap"><span class="metric-value duration-value" id="capacity-value" data-seconds="0">0 h</span></div>
            <div class="metric-meta">Man-hours/days available</div>
          </section>
          <section class="metric-card metric-card-emerald capacity-result-card">
            <div class="metric-label">Total Availability</div>
            <div class="metric-value-wrap"><span class="metric-value duration-value" id="availability-value" data-seconds="0">0 h</span></div>
            <div class="metric-meta">Capacity minus leaves</div>
          </section>
          <section class="metric-card metric-card-indigo capacity-result-card" id="capacity-tk-approved-card" hidden>
            <div class="metric-label">TK Approved (Month)</div>
            <div class="metric-value-wrap"><span class="metric-value duration-value" id="capacity-tk-approved-value" data-seconds="0">0 h</span></div>
            <div class="metric-meta" id="capacity-tk-approved-meta">Based on month filter toggle scope</div>
          </section>
          <section class="metric-card metric-card-cyan capacity-result-card" id="capacity-idle-card" hidden>
            <div class="metric-label">Idle Hours/Days (Month)</div>
            <div class="metric-value-wrap"><span class="metric-value duration-value" id="capacity-idle-value" data-seconds="0">0 h</span></div>
            <div class="metric-meta" id="capacity-idle-meta">Availability minus TK Approved for selected month</div>
          </section>
        </div>
      </div>
    </div>""")

    # Month Story Analysis panel
    parts.append("""
    <section class="panel tk-month-analysis-panel" id="tk-month-analysis" hidden>
      <div class="tk-month-analysis-header">
        <div>
          <h2>Month Story Analysis</h2>
          <div class="footnote">Stories spanning multiple months fall back to subtask estimates.</div>
        </div>
      </div>
      <div class="tk-month-chart" role="img" aria-label="Month scope estimate bar chart">
        <div class="tk-month-chart-bars">
          <section class="tk-month-chart-bar-card" data-month-analysis-slot="previous">
            <div class="tk-month-chart-value duration-value" data-month-analysis-chart-value data-seconds="0">0 h</div>
            <div class="tk-month-chart-bar-track"><div class="tk-month-chart-bar-fill" data-month-analysis-bar style="height:0%"></div></div>
            <div class="tk-month-chart-label" data-month-analysis-label>Previous Month</div>
          </section>
          <section class="tk-month-chart-bar-card featured" data-month-analysis-slot="selected">
            <div class="tk-month-chart-value duration-value" data-month-analysis-chart-value data-seconds="0">0 h</div>
            <div class="tk-month-chart-bar-track"><div class="tk-month-chart-bar-fill" data-month-analysis-bar style="height:0%"></div></div>
            <div class="tk-month-chart-label" data-month-analysis-label>Selected Month</div>
          </section>
          <section class="tk-month-chart-bar-card" data-month-analysis-slot="next">
            <div class="tk-month-chart-value duration-value" data-month-analysis-chart-value data-seconds="0">0 h</div>
            <div class="tk-month-chart-bar-track"><div class="tk-month-chart-bar-fill" data-month-analysis-bar style="height:0%"></div></div>
            <div class="tk-month-chart-label" data-month-analysis-label>Next Month</div>
          </section>
        </div>
      </div>
      <div class="tk-month-analysis-summary">
        <div class="tk-month-analysis-card">
          <div class="tk-month-analysis-card-head">
            <span>Included epics</span>
            <strong data-month-analysis-included>0</strong>
          </div>
          <div class="tk-month-analysis-epic-list" data-month-analysis-included-list></div>
        </div>
        <div class="tk-month-analysis-card excluded">
          <div class="tk-month-analysis-card-head">
            <span>Excluded epics</span>
            <strong data-month-analysis-excluded>0</strong>
          </div>
          <div class="tk-month-analysis-epic-list" data-month-analysis-excluded-list></div>
        </div>
      </div>
      <p class="footnote tk-month-analysis-scope-reconcile" data-month-analysis-scope-reconcile hidden></p>
    </section>""")

    parts.append(f"""
    <section class="panel rmi-schedule-panel" id="rmi-schedule-section">
      <div class="rmi-schedule-header-bar">
        <h2>RMI Estimation &amp; Scheduling</h2>
        <div class="rmi-schedule-controls">
          <label class="tk-month-toggle" for="rmi-jira-only-toggle">
            <input id="rmi-jira-only-toggle" class="tk-month-toggle-input" type="checkbox">
            <span class="tk-month-toggle-track" aria-hidden="true"><span class="tk-month-toggle-thumb"></span></span>
            <span class="tk-month-toggle-text">Only Jira Populated Epics</span>
          </label>
          <label for="rmi-schedule-year" class="rmi-schedule-year-label">Year</label>
          <select id="rmi-schedule-year" class="rmi-schedule-year-select" aria-label="Schedule year"></select>
        </div>
      </div>
      <div class="rmi-sched-product-cards" id="rmi-sched-product-cards"></div>
      <div class="table-frame">
        <table class="rmi-schedule-table" id="rmi-schedule-table">
          <thead>
            <tr class="rmi-sched-header-groups">
              <th></th><th></th><th></th><th></th>
              <th colspan="2" class="rmi-sched-group-estimation">Estimation</th>
              <th colspan="12" class="rmi-sched-group-scheduling">Scheduling</th>
            </tr>
            <tr class="rmi-sched-header-cols">
              <th class="rmi-sched-col-num">#</th>
              <th class="rmi-sched-col-rmi">RMI</th>
              <th class="rmi-sched-col-product">Product</th>
              <th class="rmi-sched-col-status">Status</th>
              <th class="rmi-sched-col-ml">Most&nbsp;likely</th>
              <th class="rmi-sched-col-tk">TK&nbsp;Approved</th>
              {rmi_schedule_month_headers}
            </tr>
          </thead>
          <tbody id="rmi-schedule-body"></tbody>
          <tfoot id="rmi-schedule-foot"></tfoot>
        </table>
      </div>
    </section>""")

    # Search toolbar
    parts.append("""
    <div class="search-toolbar">
      <span class="search-toolbar-label">Search</span>
      <input id="epic-search" class="search-input" type="search" placeholder="Search epics by key or title…" aria-label="Search epics">
      <button id="search-clear" class="search-clear" type="button">Clear</button>
    </div>
    <div class="search-status"></div>""")

    # View / unit / product toolbars
    parts.append(f"""
    <div class="view-toolbar">
      <button class="view-toggle active" data-view="table">Table View</button>
      <button class="view-toggle" data-view="gantt">Gantt View</button>
      <span style="width:20px"></span>
      <button class="unit-toggle active" data-unit="hours">Hours</button>
      <button class="unit-toggle" data-unit="days">Days</button>
    </div>
    <div class="product-toolbar">
      <span class="product-toolbar-label">Product</span>
      {"".join(product_toggles)}
    </div>""")

    # Table legend
    parts.append("""
    <div class="table-legend">
      <div class="table-legend-item"><span class="table-legend-swatch" style="background:var(--row-epic)"></span> Epic</div>
      <div class="table-legend-item"><span class="table-legend-swatch" style="background:var(--row-story)"></span> Story</div>
      <div class="table-legend-item"><span class="table-legend-swatch" style="background:var(--row-subtask)"></span> Subtask</div>
      <div class="table-legend-item"><span class="table-legend-swatch" style="background:var(--row-bug)"></span> Bug</div>
    </div>""")

    # Table View
    parts.append(f"""
    <section id="table-view-section">
      <div class="table-frame">
        <table class="epic-table">
          <thead><tr>
            <th style="width:42px"></th><th>Key</th><th>Summary</th><th>Product</th><th>Status</th>
            <th>Start</th><th>Due</th><th>TK Approved</th><th>Story Est.</th><th>Subtask Est.</th>
            <th>Logged</th><th>Children</th>
          </tr></thead>
          <tbody id="rmi-table-body">
            {table_body}
          </tbody>
        </table>
      </div>
    </section>""")

    # Gantt View
    parts.append("""
    <section id="gantt-view-section" hidden>
      <div id="gantt-view"></div>
    </section>""")

    # Drawer
    parts.append("""
    <div class="drawer-overlay" id="detail-drawer">
      <div class="drawer-panel">
        <div class="drawer-header">
          <h2 id="drawer-title">Details</h2>
          <button class="drawer-close" id="drawer-close">Close</button>
        </div>
        <div id="drawer-body"></div>
      </div>
    </div>""")

    # Close page div
    parts.append("</div>")

    # JSON data embed
    parts.append(f'<script type="application/json" id="rmi-report-data">{_safe_json_for_script(payload_json)}</script>')

    # JS
    parts.append("<script>" + _REPORT_JS + "</script>")
    parts.append('<script src="shared-nav.js"></script>')
    parts.append("</body>\n</html>")

    return "\n".join(parts)


def generate_html_report(db_path: Path, output_path: Path, run_id: str = "") -> Path:
    data = load_report_data(db_path, run_id)
    html = render_html(data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate database-backed RMI Jira Gantt HTML report.")
    parser.add_argument("--db", default=os.getenv("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", DEFAULT_CAPACITY_DB))
    parser.add_argument("--html", default=os.getenv("JIRA_RMI_GANTT_HTML_PATH", DEFAULT_OUTPUT_HTML))
    parser.add_argument("--run-id", default=os.getenv("JIRA_CANONICAL_RUN_ID", ""))
    parser.add_argument(
        "--canonical-db",
        default=os.getenv("JIRA_RMI_GANTT_CANONICAL_DB_PATH", ""),
        help="SQLite with canonical_issues/worklogs (default: same as --db). Env: JIRA_RMI_GANTT_CANONICAL_DB_PATH",
    )
    return parser.parse_args()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    args = parse_args()
    canon_raw = _to_text(args.canonical_db)
    if canon_raw:
        os.environ["JIRA_RMI_GANTT_CANONICAL_DB_PATH"] = str(_resolve_path(canon_raw, base_dir).resolve())
    db_path = _resolve_path(args.db, base_dir)
    output_path = _resolve_path(args.html, base_dir)
    generated = generate_html_report(db_path, output_path, args.run_id)
    print(f"Generated {generated}")


if __name__ == "__main__":
    main()
