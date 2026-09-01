"""Standalone sync for the Support Center report.

Fetches Jira issues tagged with the "Work Type EPR" custom field
(``customfield_10683``) value ``Support`` and stores their keys in a dedicated
``support_center.db``. This deliberately does NOT touch the main canonical
databases — everything else the report needs (worklogs, hierarchy, actual
completion dates, capacity) is read read-only from those existing DBs at query
time and joined back on ``issue_key``.
"""
from __future__ import annotations

import argparse
import os
import shutil
import sqlite3
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from jira_client import BASE_URL, get_session
from report_output_paths import is_writable_directory

DEFAULT_DB_NAME = "support_center.db"
DEFAULT_WORK_TYPE_FIELD_ID = "customfield_10683"
SUPPORT_VALUE = "Support"


def _base_dir() -> Path:
    return Path(__file__).resolve().parent


def _writable_db_path(candidate: Path) -> Path:
    """Keep the DB writable when the app root is a read-only Azure package mount.

    Falls back to ``$HOME/data/<name>``, seeding it once from the deployed copy so
    the packaged support data is not lost. Set ``SUPPORT_CENTER_DB_PATH`` to an
    absolute path under ``/home/data`` to control this explicitly.
    """
    if is_writable_directory(candidate.parent):
        return candidate
    fallback_dir = Path(os.getenv("HOME") or tempfile.gettempdir()) / "data"
    try:
        fallback_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        return candidate
    fallback = fallback_dir / candidate.name
    if not fallback.exists() and candidate.exists():
        try:
            shutil.copy2(candidate, fallback)
        except OSError:
            pass
    print(f"[support-center] {candidate.parent} is not writable; using {fallback}")
    return fallback


def resolve_db_path(value: str | None = None) -> Path:
    raw = (value or os.getenv("SUPPORT_CENTER_DB_PATH", DEFAULT_DB_NAME) or DEFAULT_DB_NAME).strip()
    path = Path(raw or DEFAULT_DB_NAME)
    if not path.is_absolute():
        path = _base_dir() / path
    return _writable_db_path(path)


def resolve_work_type_field_id() -> str:
    return (os.getenv("JIRA_WORK_TYPE_FIELD_ID", DEFAULT_WORK_TYPE_FIELD_ID) or DEFAULT_WORK_TYPE_FIELD_ID).strip()


def _cf_number(field_id: str) -> str:
    digits = "".join(ch for ch in str(field_id) if ch.isdigit())
    return digits


def init_support_center_db(db_path: Path) -> None:
    """Create the standalone support_center.db schema (idempotent)."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS support_issues (
                issue_key TEXT PRIMARY KEY,
                project_key TEXT NOT NULL DEFAULT '',
                issue_type TEXT NOT NULL DEFAULT '',
                summary TEXT NOT NULL DEFAULT '',
                work_type_value TEXT NOT NULL DEFAULT '',
                synced_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_support_issues_project ON support_issues(project_key)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_support_issues_type ON support_issues(issue_type)")
        conn.commit()


def _extract_option_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("value", "name"):
            candidate = str(value.get(key, "")).strip()
            if candidate:
                return candidate
        return ""
    if isinstance(value, list):
        parts = [_extract_option_value(item) for item in value]
        return ", ".join([part for part in parts if part])
    return str(value).strip()


def fetch_support_issues(session, field_id: str) -> list[dict]:
    """Page through Jira for every issue where Work Type EPR == Support."""
    cf_num = _cf_number(field_id)
    if not cf_num:
        raise ValueError(f"Could not derive a numeric custom-field id from {field_id!r}")
    jql = f'cf[{cf_num}] = "{SUPPORT_VALUE}" ORDER BY key ASC'
    url = f"{BASE_URL}/rest/api/3/search/jql"
    fields = ["project", "summary", "issuetype", field_id]
    issues: list[dict] = []
    next_page_token: str | None = None
    while True:
        payload = {"jql": jql, "maxResults": 100, "fields": fields}
        if next_page_token:
            payload["nextPageToken"] = next_page_token
        response = session.post(url, json=payload)
        response.raise_for_status()
        data = response.json()
        issues.extend(data.get("issues", []) or [])
        next_page_token = data.get("nextPageToken")
        if not next_page_token:
            break
    return issues


def _row_from_issue(issue: dict, field_id: str, synced_at: str) -> dict:
    fields = issue.get("fields", {}) or {}
    project = fields.get("project") or {}
    issue_type = fields.get("issuetype") or {}
    return {
        "issue_key": str(issue.get("key", "")).strip().upper(),
        "project_key": str(project.get("key", "")).strip().upper(),
        "issue_type": str(issue_type.get("name", "")).strip(),
        "summary": str(fields.get("summary", "") or "").strip(),
        "work_type_value": _extract_option_value(fields.get(field_id)),
        "synced_at_utc": synced_at,
    }


def write_support_issues(db_path: Path, rows: list[dict]) -> int:
    """Full-replace the support_issues table with the freshly fetched set."""
    synced = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    with sqlite3.connect(db_path) as conn:
        conn.execute("DELETE FROM support_issues")
        conn.executemany(
            """
            INSERT OR REPLACE INTO support_issues
                (issue_key, project_key, issue_type, summary, work_type_value, synced_at_utc)
            VALUES (:issue_key, :project_key, :issue_type, :summary, :work_type_value, :synced_at_utc)
            """,
            [{**row, "synced_at_utc": synced} for row in rows if row.get("issue_key")],
        )
        conn.commit()
        return conn.total_changes


def sync(db_path: Path | None = None, field_id: str | None = None) -> dict:
    target = db_path if isinstance(db_path, Path) else resolve_db_path(db_path)
    work_type_field_id = (field_id or resolve_work_type_field_id()).strip()
    init_support_center_db(target)
    session = get_session()
    issues = fetch_support_issues(session, work_type_field_id)
    synced_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rows = [_row_from_issue(issue, work_type_field_id, synced_at) for issue in issues]
    write_support_issues(target, rows)
    by_type: dict[str, int] = {}
    for row in rows:
        by_type[row["issue_type"]] = by_type.get(row["issue_type"], 0) + 1
    return {
        "db_path": str(target),
        "field_id": work_type_field_id,
        "fetched": len(rows),
        "by_type": by_type,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Sync Support-tagged Jira issues into support_center.db")
    parser.add_argument("--db", dest="db", default=None, help="Path to support_center.db")
    parser.add_argument("--field-id", dest="field_id", default=None, help="Work Type EPR custom field id (default customfield_10683)")
    args = parser.parse_args()
    result = sync(db_path=resolve_db_path(args.db) if args.db else None, field_id=args.field_id)
    print(f"support_center.db: {result['db_path']}")
    print(f"Work Type EPR field: {result['field_id']}")
    print(f"Support issues fetched: {result['fetched']}")
    for issue_type, count in sorted(result["by_type"].items()):
        print(f"  {issue_type or '(blank)'}: {count}")


if __name__ == "__main__":
    main()
