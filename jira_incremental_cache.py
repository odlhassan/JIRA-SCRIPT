from __future__ import annotations

import hashlib
import json
import os
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_iso_utc(value: str) -> datetime:
    text = str(value or "").strip()
    if not text:
        raise ValueError("empty timestamp")
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    dt = datetime.fromisoformat(text)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def get_db_path() -> Path:
    raw = os.getenv("JIRA_SYNC_DB_PATH", "jira_sync_cache.db").strip() or "jira_sync_cache.db"
    path = Path(raw)
    if path.is_absolute():
        return path
    return Path(__file__).resolve().parent / path


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS sync_state (
            pipeline TEXT PRIMARY KEY,
            last_checkpoint_utc TEXT NOT NULL,
            last_full_sync_utc TEXT,
            updated_at_utc TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS issue_index (
            issue_id TEXT PRIMARY KEY,
            issue_key TEXT NOT NULL,
            updated_utc TEXT NOT NULL,
            issue_type TEXT,
            project_key TEXT,
            last_seen_utc TEXT NOT NULL,
            is_deleted INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS issue_payloads (
            issue_id TEXT PRIMARY KEY,
            issue_key TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            updated_utc TEXT NOT NULL,
            payload_hash TEXT,
            fetched_at_utc TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS worklog_payloads (
            issue_id TEXT PRIMARY KEY,
            issue_key TEXT NOT NULL,
            worklog_json TEXT NOT NULL,
            worklog_updated_utc TEXT,
            fetched_at_utc TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS pipeline_artifacts (
            run_id TEXT,
            pipeline TEXT,
            started_at_utc TEXT,
            ended_at_utc TEXT,
            issues_scanned INTEGER,
            issues_changed INTEGER,
            new_issues INTEGER,
            detail_fetches INTEGER,
            worklog_fetches INTEGER,
            duration_ms INTEGER
        );

        CREATE INDEX IF NOT EXISTS idx_issue_index_updated ON issue_index(updated_utc);
        CREATE INDEX IF NOT EXISTS idx_issue_index_key ON issue_index(issue_key);
        """
    )
    conn.commit()


def get_or_init_checkpoint(conn: sqlite3.Connection, pipeline: str, default_utc: str) -> str:
    row = conn.execute(
        "SELECT last_checkpoint_utc FROM sync_state WHERE pipeline = ?",
        (pipeline,),
    ).fetchone()
    if row:
        return str(row[0])
    now_utc = utc_now_iso()
    conn.execute(
        """
        INSERT INTO sync_state (pipeline, last_checkpoint_utc, last_full_sync_utc, updated_at_utc)
        VALUES (?, ?, NULL, ?)
        """,
        (pipeline, default_utc, now_utc),
    )
    conn.commit()
    return default_utc


def set_checkpoint(conn: sqlite3.Connection, pipeline: str, checkpoint_utc: str) -> None:
    now_utc = utc_now_iso()
    conn.execute(
        """
        INSERT INTO sync_state (pipeline, last_checkpoint_utc, last_full_sync_utc, updated_at_utc)
        VALUES (?, ?, NULL, ?)
        ON CONFLICT(pipeline) DO UPDATE SET
            last_checkpoint_utc = excluded.last_checkpoint_utc,
            updated_at_utc = excluded.updated_at_utc
        """,
        (pipeline, checkpoint_utc, now_utc),
    )
    conn.commit()


def needs_full_sync(conn: sqlite3.Connection, pipeline: str, now_utc: str, days: int) -> bool:
    row = conn.execute(
        "SELECT last_full_sync_utc FROM sync_state WHERE pipeline = ?",
        (pipeline,),
    ).fetchone()
    if not row:
        return True
    last = str(row[0] or "").strip()
    if not last:
        return True
    return (parse_iso_utc(now_utc) - parse_iso_utc(last)) >= timedelta(days=max(int(days), 0))


def mark_full_sync(conn: sqlite3.Connection, pipeline: str, now_utc: str) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (pipeline, last_checkpoint_utc, last_full_sync_utc, updated_at_utc)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(pipeline) DO UPDATE SET
            last_full_sync_utc = excluded.last_full_sync_utc,
            updated_at_utc = excluded.updated_at_utc
        """,
        (pipeline, now_utc, now_utc, now_utc),
    )
    conn.commit()


def upsert_issue_index(conn: sqlite3.Connection, rows: list[dict]) -> None:
    if not rows:
        return
    now_utc = utc_now_iso()
    values = []
    for item in rows:
        values.append(
            (
                str(item.get("issue_id", "")).strip(),
                str(item.get("issue_key", "")).strip(),
                str(item.get("updated_utc", "")).strip(),
                str(item.get("issue_type", "")).strip(),
                str(item.get("project_key", "")).strip(),
                str(item.get("last_seen_utc", "")).strip() or now_utc,
                int(item.get("is_deleted", 0) or 0),
            )
        )
    conn.executemany(
        """
        INSERT INTO issue_index (
            issue_id, issue_key, updated_utc, issue_type, project_key, last_seen_utc, is_deleted
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(issue_id) DO UPDATE SET
            issue_key = excluded.issue_key,
            updated_utc = excluded.updated_utc,
            issue_type = excluded.issue_type,
            project_key = excluded.project_key,
            last_seen_utc = excluded.last_seen_utc,
            is_deleted = excluded.is_deleted
        """,
        values,
    )
    conn.commit()


def get_changed_or_new_issue_keys(conn: sqlite3.Connection, candidates: list[dict]) -> list[str]:
    if not candidates:
        return []
    issue_ids = [str(item.get("issue_id", "")).strip() for item in candidates if str(item.get("issue_id", "")).strip()]
    current_by_id: dict[str, tuple[str, str]] = {}
    for chunk in _chunked(issue_ids, 900):
        ph = ",".join("?" for _ in chunk)
        for row in conn.execute(
            f"SELECT issue_id, issue_key, updated_utc FROM issue_index WHERE issue_id IN ({ph})",
            tuple(chunk),
        ):
            current_by_id[str(row[0])] = (str(row[1]), str(row[2]))

    changed: list[str] = []
    seen: set[str] = set()
    for item in candidates:
        issue_id = str(item.get("issue_id", "")).strip()
        issue_key = str(item.get("issue_key", "")).strip()
        updated_utc = str(item.get("updated_utc", "")).strip()
        if not issue_id or not issue_key:
            continue
        current = current_by_id.get(issue_id)
        if current is None or current[1] != updated_utc:
            if issue_key not in seen:
                changed.append(issue_key)
                seen.add(issue_key)
    return changed


def upsert_issue_payloads(conn: sqlite3.Connection, payloads: list[dict]) -> None:
    if not payloads:
        return
    now_utc = utc_now_iso()
    values = []
    for item in payloads:
        issue_id = str(item.get("issue_id", "")).strip()
        issue_key = str(item.get("issue_key", "")).strip()
        updated_utc = str(item.get("updated_utc", "")).strip()
        payload = item.get("payload", {})
        payload_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        payload_hash = hashlib.sha256(payload_json.encode("utf-8")).hexdigest()
        values.append((issue_id, issue_key, payload_json, updated_utc, payload_hash, now_utc))
    conn.executemany(
        """
        INSERT INTO issue_payloads (
            issue_id, issue_key, payload_json, updated_utc, payload_hash, fetched_at_utc
        )
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(issue_id) DO UPDATE SET
            issue_key = excluded.issue_key,
            payload_json = excluded.payload_json,
            updated_utc = excluded.updated_utc,
            payload_hash = excluded.payload_hash,
            fetched_at_utc = excluded.fetched_at_utc
        """,
        values,
    )
    conn.commit()


def get_cached_issue_payloads(
    conn: sqlite3.Connection,
    project_keys: list[str],
    issue_types: list[str],
) -> list[dict]:
    project_values = [str(v).strip() for v in project_keys if str(v).strip()]
    issue_type_values = [str(v).strip().lower() for v in issue_types if str(v).strip()]

    where = ["idx.is_deleted = 0"]
    params: list[str] = []
    if project_values:
        where.append(f"idx.project_key IN ({','.join('?' for _ in project_values)})")
        params.extend(project_values)
    if issue_type_values:
        where.append(f"LOWER(idx.issue_type) IN ({','.join('?' for _ in issue_type_values)})")
        params.extend(issue_type_values)

    query = (
        "SELECT p.payload_json FROM issue_payloads p "
        "JOIN issue_index idx ON idx.issue_id = p.issue_id "
        f"WHERE {' AND '.join(where)}"
    )
    rows = conn.execute(query, tuple(params)).fetchall()
    results: list[dict] = []
    for row in rows:
        try:
            payload = json.loads(str(row[0]))
            if isinstance(payload, dict):
                results.append(payload)
        except Exception:
            continue
    return results


def upsert_worklog_payload(
    conn: sqlite3.Connection,
    issue_key: str,
    issue_id: str,
    worklogs: list[dict],
    worklog_updated_utc: str | None,
) -> None:
    now_utc = utc_now_iso()
    worklog_json = json.dumps(worklogs or [], ensure_ascii=False, separators=(",", ":"))
    conn.execute(
        """
        INSERT INTO worklog_payloads (
            issue_id, issue_key, worklog_json, worklog_updated_utc, fetched_at_utc
        )
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(issue_id) DO UPDATE SET
            issue_key = excluded.issue_key,
            worklog_json = excluded.worklog_json,
            worklog_updated_utc = excluded.worklog_updated_utc,
            fetched_at_utc = excluded.fetched_at_utc
        """,
        (issue_id, issue_key, worklog_json, worklog_updated_utc, now_utc),
    )
    conn.commit()


def get_cached_worklogs_for_subtasks(
    conn: sqlite3.Connection,
    issue_keys: list[str],
) -> dict[str, list[dict]]:
    keys = [str(v).strip() for v in issue_keys if str(v).strip()]
    if not keys:
        return {}
    results: dict[str, list[dict]] = {}
    for chunk in _chunked(keys, 900):
        ph = ",".join("?" for _ in chunk)
        for row in conn.execute(
            f"SELECT issue_key, worklog_json FROM worklog_payloads WHERE issue_key IN ({ph})",
            tuple(chunk),
        ):
            issue_key = str(row[0])
            try:
                payload = json.loads(str(row[1]))
                results[issue_key] = payload if isinstance(payload, list) else []
            except Exception:
                results[issue_key] = []
    return results


def mark_missing_issues_deleted(
    conn: sqlite3.Connection,
    project_keys: list[str],
    issue_types: list[str],
    active_issue_ids: set[str],
) -> int:
    projects = [str(v).strip() for v in project_keys if str(v).strip()]
    kinds = [str(v).strip().lower() for v in issue_types if str(v).strip()]
    where = ["is_deleted = 0"]
    params: list[str] = []
    if projects:
        where.append(f"project_key IN ({','.join('?' for _ in projects)})")
        params.extend(projects)
    if kinds:
        where.append(f"LOWER(issue_type) IN ({','.join('?' for _ in kinds)})")
        params.extend(kinds)

    candidates = [
        str(row[0])
        for row in conn.execute(
            f"SELECT issue_id FROM issue_index WHERE {' AND '.join(where)}",
            tuple(params),
        ).fetchall()
    ]
    stale = [issue_id for issue_id in candidates if issue_id not in active_issue_ids]
    if not stale:
        return 0
    conn.executemany(
        "UPDATE issue_index SET is_deleted = 1 WHERE issue_id = ?",
        [(issue_id,) for issue_id in stale],
    )
    conn.commit()
    return len(stale)


def record_pipeline_artifact(conn: sqlite3.Connection, payload: dict) -> None:
    conn.execute(
        """
        INSERT INTO pipeline_artifacts (
            run_id, pipeline, started_at_utc, ended_at_utc,
            issues_scanned, issues_changed, new_issues,
            detail_fetches, worklog_fetches, duration_ms
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload.get("run_id"),
            payload.get("pipeline"),
            payload.get("started_at_utc"),
            payload.get("ended_at_utc"),
            int(payload.get("issues_scanned") or 0),
            int(payload.get("issues_changed") or 0),
            int(payload.get("new_issues") or 0),
            int(payload.get("detail_fetches") or 0),
            int(payload.get("worklog_fetches") or 0),
            int(payload.get("duration_ms") or 0),
        ),
    )
    conn.commit()


def bootstrap_default_checkpoint(days: int) -> str:
    age = max(int(days), 0)
    return (datetime.now(timezone.utc) - timedelta(days=age)).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def apply_overlap(checkpoint_utc: str, overlap_minutes: int) -> str:
    dt = parse_iso_utc(checkpoint_utc) - timedelta(minutes=max(int(overlap_minutes), 0))
    return dt.replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _chunked(values: Iterable[str], size: int):
    chunk: list[str] = []
    for value in values:
        chunk.append(value)
        if len(chunk) >= size:
            yield chunk
            chunk = []
    if chunk:
        yield chunk
