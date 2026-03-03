from __future__ import annotations

import csv
import hashlib
import io
import json
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook


VALID_MODES = {"log_date", "planned_dates"}
DEFAULT_RETENTION_DAYS = 30
RUN_STATUSES = {"queued", "running", "success", "failed", "canceled"}


@dataclass(frozen=True)
class SnapshotFilter:
    from_date: str
    to_date: str
    mode: str
    projects_scope: str
    statuses_scope: str
    assignees_scope: str



def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()



def _norm_set(values: set[str], *, upper: bool = False) -> list[str]:
    out = set()
    for item in values or set():
        txt = _text(item)
        if not txt:
            continue
        out.add(txt.upper() if upper else txt.lower())
    return sorted(out)



def _scope(values: set[str], *, upper: bool = False) -> str:
    return ",".join(_norm_set(values, upper=upper))



def make_filter(
    from_date: str,
    to_date: str,
    mode: str,
    selected_projects: set[str],
    selected_statuses: set[str],
    selected_assignees: set[str],
) -> SnapshotFilter:
    return SnapshotFilter(
        from_date=_text(from_date),
        to_date=_text(to_date),
        mode=_text(mode).lower() or "log_date",
        projects_scope=_scope(selected_projects, upper=True),
        statuses_scope=_scope(selected_statuses, upper=False),
        assignees_scope=_scope(selected_assignees, upper=False),
    )



def init_db(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_refresh_runs (
                run_id TEXT PRIMARY KEY,
                from_date TEXT NOT NULL,
                to_date TEXT NOT NULL,
                mode TEXT NOT NULL,
                projects_scope TEXT NOT NULL DEFAULT '',
                statuses_scope TEXT NOT NULL DEFAULT '',
                assignees_scope TEXT NOT NULL DEFAULT '',
                force_full INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL,
                progress_step TEXT NOT NULL DEFAULT '',
                progress_pct INTEGER NOT NULL DEFAULT 0,
                source TEXT NOT NULL DEFAULT '',
                started_at_utc TEXT NOT NULL,
                completed_at_utc TEXT NOT NULL DEFAULT '',
                error_text TEXT NOT NULL DEFAULT '',
                stats_json TEXT NOT NULL DEFAULT '{}',
                cancel_requested INTEGER NOT NULL DEFAULT 0,
                queued_at_utc TEXT NOT NULL DEFAULT '',
                attempt INTEGER NOT NULL DEFAULT 1,
                max_attempts INTEGER NOT NULL DEFAULT 1,
                next_retry_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        cols = conn.execute("PRAGMA table_info(planned_actual_refresh_runs)").fetchall()
        col_names = {str(item[1]) for item in cols}
        if "cancel_requested" not in col_names:
            conn.execute("ALTER TABLE planned_actual_refresh_runs ADD COLUMN cancel_requested INTEGER NOT NULL DEFAULT 0")
        if "queued_at_utc" not in col_names:
            conn.execute("ALTER TABLE planned_actual_refresh_runs ADD COLUMN queued_at_utc TEXT NOT NULL DEFAULT ''")
        if "attempt" not in col_names:
            conn.execute("ALTER TABLE planned_actual_refresh_runs ADD COLUMN attempt INTEGER NOT NULL DEFAULT 1")
        if "max_attempts" not in col_names:
            conn.execute("ALTER TABLE planned_actual_refresh_runs ADD COLUMN max_attempts INTEGER NOT NULL DEFAULT 1")
        if "next_retry_at_utc" not in col_names:
            conn.execute("ALTER TABLE planned_actual_refresh_runs ADD COLUMN next_retry_at_utc TEXT NOT NULL DEFAULT ''")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_snapshots (
                snapshot_id TEXT PRIMARY KEY,
                from_date TEXT NOT NULL,
                to_date TEXT NOT NULL,
                mode TEXT NOT NULL,
                projects_scope TEXT NOT NULL DEFAULT '',
                statuses_scope TEXT NOT NULL DEFAULT '',
                assignees_scope TEXT NOT NULL DEFAULT '',
                rows_json TEXT NOT NULL,
                totals_json TEXT NOT NULL,
                options_json TEXT NOT NULL DEFAULT '{}',
                source_json TEXT NOT NULL DEFAULT '{}',
                watermark_utc TEXT NOT NULL DEFAULT '',
                computed_at_utc TEXT NOT NULL,
                row_count INTEGER NOT NULL DEFAULT 0,
                is_official INTEGER NOT NULL DEFAULT 0,
                official_pinned_by TEXT NOT NULL DEFAULT '',
                official_pinned_at_utc TEXT NOT NULL DEFAULT '',
                lifecycle_state TEXT NOT NULL DEFAULT 'active'
            )
            """
        )
        snap_cols = conn.execute("PRAGMA table_info(planned_actual_snapshots)").fetchall()
        snap_col_names = {str(item[1]) for item in snap_cols}
        if "is_official" not in snap_col_names:
            conn.execute("ALTER TABLE planned_actual_snapshots ADD COLUMN is_official INTEGER NOT NULL DEFAULT 0")
        if "official_pinned_by" not in snap_col_names:
            conn.execute("ALTER TABLE planned_actual_snapshots ADD COLUMN official_pinned_by TEXT NOT NULL DEFAULT ''")
        if "official_pinned_at_utc" not in snap_col_names:
            conn.execute("ALTER TABLE planned_actual_snapshots ADD COLUMN official_pinned_at_utc TEXT NOT NULL DEFAULT ''")
        if "lifecycle_state" not in snap_col_names:
            conn.execute("ALTER TABLE planned_actual_snapshots ADD COLUMN lifecycle_state TEXT NOT NULL DEFAULT 'active'")
        conn.execute(
            """
            DROP INDEX IF EXISTS idx_planned_actual_snapshot_filter
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_planned_actual_snapshot_filter
            ON planned_actual_snapshots(from_date, to_date, mode, projects_scope, statuses_scope, assignees_scope)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_planned_actual_snapshot_computed
            ON planned_actual_snapshots(computed_at_utc DESC)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_sync_state (
                scope_key TEXT PRIMARY KEY,
                last_successful_fetch_utc TEXT NOT NULL,
                last_run_id TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_source_audit (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id TEXT NOT NULL,
                snapshot_id TEXT NOT NULL,
                source_type TEXT NOT NULL,
                counts_json TEXT NOT NULL DEFAULT '{}',
                payload_checksum TEXT NOT NULL DEFAULT '',
                created_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_ui_settings (
                id INTEGER PRIMARY KEY CHECK (id=1),
                settings_json TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS planned_actual_snapshot_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_id TEXT NOT NULL,
                run_id TEXT NOT NULL DEFAULT '',
                event_type TEXT NOT NULL,
                actor TEXT NOT NULL DEFAULT '',
                details_json TEXT NOT NULL DEFAULT '{}',
                created_at_utc TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()



def create_run(
    db_path: Path,
    run_id: str,
    flt: SnapshotFilter,
    force_full: bool,
    *,
    status: str = "running",
    progress_step: str = "",
    progress_pct: int = 0,
    attempt: int = 1,
    max_attempts: int = 1,
    next_retry_at_utc: str = "",
) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    normalized_status = _text(status).lower() or "running"
    if normalized_status not in RUN_STATUSES:
        normalized_status = "running"
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_refresh_runs(
                run_id, from_date, to_date, mode, projects_scope, statuses_scope, assignees_scope,
                force_full, status, progress_step, progress_pct, started_at_utc, queued_at_utc,
                attempt, max_attempts, next_retry_at_utc
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                flt.from_date,
                flt.to_date,
                flt.mode,
                flt.projects_scope,
                flt.statuses_scope,
                flt.assignees_scope,
                1 if force_full else 0,
                normalized_status,
                _text(progress_step),
                int(max(0, min(100, progress_pct))),
                now,
                now if normalized_status == "queued" else "",
                max(1, int(attempt or 1)),
                max(1, int(max_attempts or 1)),
                _text(next_retry_at_utc),
            ),
        )
        conn.commit()
    finally:
        conn.close()



def update_run_progress(db_path: Path, run_id: str, step: str, pct: int) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "UPDATE planned_actual_refresh_runs SET progress_step=?, progress_pct=? WHERE run_id=?",
            (_text(step), int(max(0, min(100, pct))), run_id),
        )
        conn.commit()
    finally:
        conn.close()



def finish_run_success(db_path: Path, run_id: str, source: str, stats: dict[str, Any]) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE planned_actual_refresh_runs
            SET status='success', progress_step='completed', progress_pct=100,
                completed_at_utc=?, source=?, stats_json=?, cancel_requested=0
            WHERE run_id=?
            """,
            (now, _text(source), json.dumps(stats or {}, ensure_ascii=True), run_id),
        )
        conn.commit()
    finally:
        conn.close()



def finish_run_failed(db_path: Path, run_id: str, error_text: str, stats: dict[str, Any] | None = None) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE planned_actual_refresh_runs
            SET status='failed', completed_at_utc=?, error_text=?, stats_json=?, cancel_requested=0
            WHERE run_id=?
            """,
            (now, _text(error_text), json.dumps(stats or {}, ensure_ascii=True), run_id),
        )
        conn.commit()
    finally:
        conn.close()


def request_cancel(db_path: Path, run_id: str) -> bool:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            """
            UPDATE planned_actual_refresh_runs
            SET cancel_requested=1,
                progress_step=CASE WHEN status='queued' THEN 'canceling' ELSE progress_step END
            WHERE run_id=? AND status IN ('queued','running')
            """,
            (_text(run_id),),
        )
        conn.commit()
        return int(cur.rowcount or 0) > 0
    finally:
        conn.close()


def is_cancel_requested(db_path: Path, run_id: str) -> bool:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT cancel_requested FROM planned_actual_refresh_runs WHERE run_id=?",
            (_text(run_id),),
        ).fetchone()
        return bool(int((row["cancel_requested"] if row else 0) or 0))
    finally:
        conn.close()


def mark_run_canceled(db_path: Path, run_id: str, reason: str = "") -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE planned_actual_refresh_runs
            SET status='canceled',
                progress_step='canceled',
                progress_pct=0,
                completed_at_utc=?,
                error_text=?,
                cancel_requested=0
            WHERE run_id=?
            """,
            (now, _text(reason) or "Canceled by user (rollback applied).", _text(run_id)),
        )
        conn.commit()
    finally:
        conn.close()


def begin_queued_run(db_path: Path, run_id: str) -> bool:
    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT status, cancel_requested FROM planned_actual_refresh_runs WHERE run_id=?",
            (_text(run_id),),
        ).fetchone()
        if not row:
            return False
        status = _text(row[0]).lower()
        cancel_requested = int(row[1] or 0)
        if status != "queued":
            return False
        if cancel_requested:
            return False
        conn.execute(
            """
            UPDATE planned_actual_refresh_runs
            SET status='running', progress_step='initializing', progress_pct=1, queued_at_utc=''
            WHERE run_id=?
            """,
            (_text(run_id),),
        )
        conn.commit()
        return True
    finally:
        conn.close()



def get_run(db_path: Path, run_id: str) -> dict[str, Any] | None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT * FROM planned_actual_refresh_runs WHERE run_id=?",
            (run_id,),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()



def has_running_run_for_scope(db_path: Path, flt: SnapshotFilter) -> bool:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT run_id
            FROM planned_actual_refresh_runs
            WHERE status='running' AND from_date=? AND to_date=? AND mode=?
              AND projects_scope=? AND statuses_scope=? AND assignees_scope=?
            LIMIT 1
            """,
            (
                flt.from_date,
                flt.to_date,
                flt.mode,
                flt.projects_scope,
                flt.statuses_scope,
                flt.assignees_scope,
            ),
        ).fetchone()
        return row is not None
    finally:
        conn.close()



def save_snapshot(
    db_path: Path,
    snapshot_id: str,
    flt: SnapshotFilter,
    rows: list[dict[str, Any]],
    totals: dict[str, Any],
    options: dict[str, Any],
    source: dict[str, Any],
    watermark_utc: str,
) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_snapshots(
                snapshot_id, from_date, to_date, mode, projects_scope, statuses_scope, assignees_scope,
                rows_json, totals_json, options_json, source_json, watermark_utc, computed_at_utc, row_count
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                snapshot_id,
                flt.from_date,
                flt.to_date,
                flt.mode,
                flt.projects_scope,
                flt.statuses_scope,
                flt.assignees_scope,
                json.dumps(rows or [], ensure_ascii=True, separators=(",", ":")),
                json.dumps(totals or {}, ensure_ascii=True, separators=(",", ":")),
                json.dumps(options or {}, ensure_ascii=True, separators=(",", ":")),
                json.dumps(source or {}, ensure_ascii=True, separators=(",", ":")),
                _text(watermark_utc),
                now,
                len(rows or []),
            ),
        )
        conn.commit()
    finally:
        conn.close()



def _row_to_snapshot_payload(row: sqlite3.Row) -> dict[str, Any]:
    payload = {
        "snapshot_id": _text(row["snapshot_id"]),
        "from_date": _text(row["from_date"]),
        "to_date": _text(row["to_date"]),
        "mode": _text(row["mode"]),
        "projects_scope": _text(row["projects_scope"]),
        "statuses_scope": _text(row["statuses_scope"]),
        "assignees_scope": _text(row["assignees_scope"]),
        "computed_at_utc": _text(row["computed_at_utc"]),
        "watermark_utc": _text(row["watermark_utc"]),
        "is_official": bool(int(row["is_official"] or 0)),
        "official_pinned_by": _text(row["official_pinned_by"]),
        "official_pinned_at_utc": _text(row["official_pinned_at_utc"]),
        "lifecycle_state": _text(row["lifecycle_state"]) or "active",
        "rows": json.loads(_text(row["rows_json"]) or "[]"),
        "totals": json.loads(_text(row["totals_json"]) or "{}"),
        "options": json.loads(_text(row["options_json"]) or "{}"),
        "source": json.loads(_text(row["source_json"]) or "{}"),
    }
    return payload



def load_snapshot_by_filter(db_path: Path, flt: SnapshotFilter) -> dict[str, Any] | None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT *
            FROM planned_actual_snapshots
            WHERE from_date=? AND to_date=? AND mode=?
              AND projects_scope=? AND statuses_scope=? AND assignees_scope=?
            ORDER BY computed_at_utc DESC, rowid DESC
            LIMIT 1
            """,
            (
                flt.from_date,
                flt.to_date,
                flt.mode,
                flt.projects_scope,
                flt.statuses_scope,
                flt.assignees_scope,
            ),
        ).fetchone()
        return _row_to_snapshot_payload(row) if row else None
    finally:
        conn.close()



def load_snapshot_by_id(db_path: Path, snapshot_id: str) -> dict[str, Any] | None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT * FROM planned_actual_snapshots WHERE snapshot_id=?",
            (_text(snapshot_id),),
        ).fetchone()
        return _row_to_snapshot_payload(row) if row else None
    finally:
        conn.close()



def list_history(db_path: Path, limit: int = 30) -> list[dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT s.snapshot_id, s.from_date, s.to_date, s.mode, s.projects_scope, s.statuses_scope,
                   s.assignees_scope, s.computed_at_utc, s.row_count, s.is_official, s.official_pinned_by, s.official_pinned_at_utc, s.lifecycle_state,
                   r.run_id, r.status, r.source
            FROM planned_actual_snapshots s
            LEFT JOIN planned_actual_refresh_runs r ON r.run_id = (
              SELECT rr.run_id FROM planned_actual_refresh_runs rr
              WHERE rr.from_date=s.from_date AND rr.to_date=s.to_date AND rr.mode=s.mode
                AND rr.projects_scope=s.projects_scope AND rr.statuses_scope=s.statuses_scope AND rr.assignees_scope=s.assignees_scope
                AND rr.status='success'
              ORDER BY rr.completed_at_utc DESC LIMIT 1
            )
            ORDER BY s.computed_at_utc DESC
            LIMIT ?
            """,
            (max(1, min(500, int(limit))),),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def list_queue(db_path: Path, limit: int = 50) -> list[dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT run_id, from_date, to_date, mode, projects_scope, statuses_scope, assignees_scope,
                   force_full, status, progress_step, progress_pct, source,
                   started_at_utc, completed_at_utc, error_text, cancel_requested,
                   queued_at_utc, attempt, max_attempts, next_retry_at_utc
            FROM planned_actual_refresh_runs
            WHERE status IN ('queued', 'running')
            ORDER BY
              CASE status WHEN 'running' THEN 0 ELSE 1 END ASC,
              started_at_utc ASC
            LIMIT ?
            """,
            (max(1, min(500, int(limit))),),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()



def save_sync_state(db_path: Path, scope_key: str, last_fetch_utc: str, run_id: str) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_sync_state(scope_key, last_successful_fetch_utc, last_run_id, updated_at_utc)
            VALUES(?, ?, ?, ?)
            ON CONFLICT(scope_key) DO UPDATE SET
                last_successful_fetch_utc=excluded.last_successful_fetch_utc,
                last_run_id=excluded.last_run_id,
                updated_at_utc=excluded.updated_at_utc
            """,
            (_text(scope_key), _text(last_fetch_utc), _text(run_id), now),
        )
        conn.commit()
    finally:
        conn.close()



def load_sync_state(db_path: Path, scope_key: str) -> dict[str, Any] | None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT * FROM planned_actual_sync_state WHERE scope_key=?",
            (_text(scope_key),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()



def save_source_audit(
    db_path: Path,
    run_id: str,
    snapshot_id: str,
    source_type: str,
    counts: dict[str, Any],
    payload: dict[str, Any],
) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    checksum = hashlib.sha1(json.dumps(payload or {}, sort_keys=True, ensure_ascii=True).encode("utf-8")).hexdigest()
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_source_audit(
                run_id, snapshot_id, source_type, counts_json, payload_checksum, created_at_utc
            ) VALUES(?, ?, ?, ?, ?, ?)
            """,
            (
                _text(run_id),
                _text(snapshot_id),
                _text(source_type),
                json.dumps(counts or {}, ensure_ascii=True),
                checksum,
                now,
            ),
        )
        conn.commit()
    finally:
        conn.close()



def prune_old_data(db_path: Path, retention_days: int = DEFAULT_RETENTION_DAYS) -> None:
    days = max(1, int(retention_days or DEFAULT_RETENTION_DAYS))
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    cutoff_text = cutoff.strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("DELETE FROM planned_actual_snapshots WHERE computed_at_utc < ?", (cutoff_text,))
        conn.execute("DELETE FROM planned_actual_refresh_runs WHERE started_at_utc < ?", (cutoff_text,))
        conn.execute("DELETE FROM planned_actual_source_audit WHERE created_at_utc < ?", (cutoff_text,))
        conn.execute("DELETE FROM planned_actual_snapshot_events WHERE created_at_utc < ?", (cutoff_text,))
        conn.commit()
    finally:
        conn.close()


def save_snapshot_event(
    db_path: Path,
    snapshot_id: str,
    event_type: str,
    actor: str,
    details: dict[str, Any] | None = None,
    run_id: str = "",
) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_snapshot_events(snapshot_id, run_id, event_type, actor, details_json, created_at_utc)
            VALUES(?, ?, ?, ?, ?, ?)
            """,
            (
                _text(snapshot_id),
                _text(run_id),
                _text(event_type),
                _text(actor),
                json.dumps(details or {}, ensure_ascii=True, separators=(",", ":")),
                now,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def pin_official_snapshot(db_path: Path, snapshot_id: str, actor: str) -> dict[str, Any] | None:
    snapshot = load_snapshot_by_id(db_path, snapshot_id)
    if not snapshot:
        return None
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE planned_actual_snapshots
            SET is_official=0, official_pinned_by='', official_pinned_at_utc=''
            WHERE from_date=? AND to_date=? AND mode=? AND projects_scope=? AND statuses_scope=? AND assignees_scope=?
            """,
            (
                _text(snapshot.get("from_date")),
                _text(snapshot.get("to_date")),
                _text(snapshot.get("mode")),
                _text(snapshot.get("projects_scope")),
                _text(snapshot.get("statuses_scope")),
                _text(snapshot.get("assignees_scope")),
            ),
        )
        conn.execute(
            """
            UPDATE planned_actual_snapshots
            SET is_official=1, official_pinned_by=?, official_pinned_at_utc=?
            WHERE snapshot_id=?
            """,
            (_text(actor), now, _text(snapshot_id)),
        )
        conn.commit()
    finally:
        conn.close()
    save_snapshot_event(
        db_path=db_path,
        snapshot_id=snapshot_id,
        event_type="pin_official",
        actor=actor,
        details={"pinned_at_utc": now},
    )
    return load_snapshot_by_id(db_path, snapshot_id)


def unpin_official_snapshot(db_path: Path, snapshot_id: str, actor: str) -> dict[str, Any] | None:
    snapshot = load_snapshot_by_id(db_path, snapshot_id)
    if not snapshot:
        return None
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE planned_actual_snapshots
            SET is_official=0, official_pinned_by='', official_pinned_at_utc=''
            WHERE snapshot_id=?
            """,
            (_text(snapshot_id),),
        )
        conn.commit()
    finally:
        conn.close()
    save_snapshot_event(
        db_path=db_path,
        snapshot_id=snapshot_id,
        event_type="unpin_official",
        actor=actor,
        details={},
    )
    return load_snapshot_by_id(db_path, snapshot_id)



def load_ui_settings(db_path: Path) -> dict[str, Any]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute("SELECT settings_json FROM planned_actual_ui_settings WHERE id=1").fetchone()
        if not row:
            return {"column_widths": {}, "density": "comfortable", "sort": {"key": "", "direction": "asc"}}
        try:
            parsed = json.loads(_text(row["settings_json"]) or "{}")
            return parsed if isinstance(parsed, dict) else {}
        except Exception:
            return {}
    finally:
        conn.close()



def save_ui_settings(db_path: Path, settings: dict[str, Any]) -> dict[str, Any]:
    payload = settings if isinstance(settings, dict) else {}
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO planned_actual_ui_settings(id, settings_json, updated_at_utc)
            VALUES(1, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                settings_json=excluded.settings_json,
                updated_at_utc=excluded.updated_at_utc
            """,
            (json.dumps(payload, ensure_ascii=True), now),
        )
        conn.commit()
    finally:
        conn.close()
    return payload



def build_snapshot_payload(
    hierarchy: dict[str, Any],
    actual_hours_by_subtask: dict[str, float],
    selected_projects: set[str],
    selected_statuses: set[str],
    selected_assignees: set[str],
) -> tuple[list[dict[str, Any]], dict[str, Any], dict[str, Any]]:
    epics = hierarchy.get("epics", []) or []
    stories = hierarchy.get("stories", []) or []
    subtasks = hierarchy.get("subtasks", []) or []

    statuses_norm = {item.lower() for item in _norm_set(selected_statuses, upper=False)}
    assignees_norm = {item.lower() for item in _norm_set(selected_assignees, upper=False)}
    projects_norm = {item.upper() for item in _norm_set(selected_projects, upper=True)}

    def _include(project_key: str, status: str, assignee: str) -> bool:
        if projects_norm and _text(project_key).upper() not in projects_norm:
            return False
        if statuses_norm and _text(status).lower() not in statuses_norm:
            return False
        if assignees_norm and _text(assignee).lower() not in assignees_norm:
            return False
        return True

    stories_by_epic: dict[str, list[dict[str, Any]]] = {}
    for story in stories:
        epic_key = _text(story.get("epic_key")).upper()
        stories_by_epic.setdefault(epic_key, []).append(story)

    subtasks_by_story: dict[str, list[dict[str, Any]]] = {}
    for sub in subtasks:
        story_key = _text(sub.get("story_key")).upper()
        subtasks_by_story.setdefault(story_key, []).append(sub)

    rows: list[dict[str, Any]] = []
    project_rollup: dict[str, dict[str, Any]] = {}

    def _with_display_metrics(base: dict[str, Any]) -> dict[str, Any]:
        man_hours = float(base.get("planned_hours") or 0.0)
        actual_hours = float(base.get("actual_hours") or 0.0)
        delta_hours = float(base.get("variance_hours") or (man_hours - actual_hours))
        base["man_hours"] = round(man_hours, 2)
        base["actual_hours"] = round(actual_hours, 2)
        base["delta_hours"] = round(delta_hours, 2)
        base["man_days"] = round(man_hours / 8.0, 2)
        base["actual_days"] = round(actual_hours / 8.0, 2)
        base["delta_days"] = round(delta_hours / 8.0, 2)
        base["planned_start_date"] = _text(base.get("planned_start"))
        base["planned_end_date"] = _text(base.get("planned_due"))
        base["aspect"] = _text(base.get("summary"))
        base["type"] = _text(base.get("row_type"))
        base["resource_logged"] = _text(base.get("assignee")) if _text(base.get("assignee")) else "N/A"
        return base

    for epic in epics:
        project_key = _text(epic.get("project_key")).upper() or "UNKNOWN"
        project_name = _text(epic.get("project_name")) or project_key
        epic_status = _text(epic.get("status"))
        epic_assignee = _text(epic.get("assignee")) or "Unassigned"
        if not _include(project_key, epic_status, epic_assignee):
            continue

        epic_key = _text(epic.get("issue_key")).upper()
        epic_planned = float(epic.get("estimate_hours") or 0.0)
        epic_actual = 0.0
        story_count = 0
        subtask_count = 0

        story_rows: list[dict[str, Any]] = []
        for story in sorted(stories_by_epic.get(epic_key, []), key=lambda x: (_text(x.get("summary")).lower(), _text(x.get("issue_key")))):
            story_status = _text(story.get("status"))
            story_assignee = _text(story.get("assignee")) or "Unassigned"
            if not _include(project_key, story_status, story_assignee):
                continue
            story_key = _text(story.get("issue_key")).upper()
            story_planned = float(story.get("estimate_hours") or 0.0)
            story_actual = 0.0
            filtered_subtasks = []
            for sub in sorted(subtasks_by_story.get(story_key, []), key=lambda x: (_text(x.get("summary")).lower(), _text(x.get("issue_key")))):
                sub_status = _text(sub.get("status"))
                sub_assignee = _text(sub.get("assignee")) or "Unassigned"
                if not _include(project_key, sub_status, sub_assignee):
                    continue
                sub_key = _text(sub.get("issue_key")).upper()
                sub_planned = float(sub.get("estimate_hours") or 0.0)
                sub_actual = float(actual_hours_by_subtask.get(sub_key, 0.0) or 0.0)
                story_actual += sub_actual
                filtered_subtasks.append(
                    _with_display_metrics({
                        "row_type": "subtask",
                        "project_key": project_key,
                        "project_name": project_name,
                        "issue_key": sub_key,
                        "parent_key": story_key,
                        "summary": _text(sub.get("summary")) or sub_key,
                        "assignee": sub_assignee,
                        "status": sub_status,
                        "planned_start": _text(sub.get("planned_start")),
                        "planned_due": _text(sub.get("planned_due")),
                        "planned_hours": round(sub_planned, 2),
                        "actual_hours": round(sub_actual, 2),
                        "variance_hours": round(sub_planned - sub_actual, 2),
                    })
                )

            story_count += 1
            subtask_count += len(filtered_subtasks)
            epic_actual += story_actual
            story_rows.append(
                _with_display_metrics({
                    "row_type": "story",
                    "project_key": project_key,
                    "project_name": project_name,
                    "issue_key": story_key,
                    "parent_key": epic_key,
                    "summary": _text(story.get("summary")) or story_key,
                    "assignee": story_assignee,
                    "status": story_status,
                    "planned_start": _text(story.get("planned_start")),
                    "planned_due": _text(story.get("planned_due")),
                    "planned_hours": round(story_planned, 2),
                    "actual_hours": round(story_actual, 2),
                    "variance_hours": round(story_planned - story_actual, 2),
                    "children": filtered_subtasks,
                })
            )

        project_rollup.setdefault(
            project_key,
            {
                "project_key": project_key,
                "project_name": project_name,
                "planned_hours": 0.0,
                "actual_hours": 0.0,
                "epic_count": 0,
                "story_count": 0,
                "subtask_count": 0,
            },
        )
        project_rollup[project_key]["planned_hours"] += epic_planned
        project_rollup[project_key]["actual_hours"] += epic_actual
        project_rollup[project_key]["epic_count"] += 1
        project_rollup[project_key]["story_count"] += story_count
        project_rollup[project_key]["subtask_count"] += subtask_count

        rows.append(
            _with_display_metrics({
                "row_type": "epic",
                "project_key": project_key,
                "project_name": project_name,
                "issue_key": epic_key,
                "parent_key": project_key,
                "summary": _text(epic.get("summary")) or epic_key,
                "assignee": epic_assignee,
                "status": epic_status,
                "planned_start": _text(epic.get("planned_start")),
                "planned_due": _text(epic.get("planned_due")),
                "planned_hours": round(epic_planned, 2),
                "actual_hours": round(epic_actual, 2),
                "variance_hours": round(epic_planned - epic_actual, 2),
                "story_count": story_count,
                "subtask_count": subtask_count,
                "children": story_rows,
            })
        )

    project_rows = []
    for key in sorted(project_rollup):
        item = project_rollup[key]
        project_rows.append(
            _with_display_metrics({
                "row_type": "project",
                "project_key": key,
                "project_name": _text(item.get("project_name")) or key,
                "issue_key": key,
                "parent_key": "",
                "summary": _text(item.get("project_name")) or key,
                "assignee": "",
                "status": "",
                "planned_start": "",
                "planned_due": "",
                "planned_hours": round(float(item.get("planned_hours") or 0.0), 2),
                "actual_hours": round(float(item.get("actual_hours") or 0.0), 2),
                "variance_hours": round(float(item.get("planned_hours") or 0.0) - float(item.get("actual_hours") or 0.0), 2),
                "epic_count": int(item.get("epic_count") or 0),
                "story_count": int(item.get("story_count") or 0),
                "subtask_count": int(item.get("subtask_count") or 0),
            })
        )

    rows_sorted = sorted(rows, key=lambda x: (_text(x.get("project_key")), _text(x.get("summary")).lower(), _text(x.get("issue_key"))))
    final_rows = project_rows + rows_sorted

    total_planned = round(sum(float(item.get("planned_hours") or 0.0) for item in project_rows), 2)
    total_actual = round(sum(float(item.get("actual_hours") or 0.0) for item in project_rows), 2)
    totals = {
        "planned_hours": total_planned,
        "actual_hours": total_actual,
        "variance_hours": round(total_planned - total_actual, 2),
        "project_count": len(project_rows),
        "epic_count": sum(int(item.get("epic_count") or 0) for item in project_rows),
        "story_count": sum(int(item.get("story_count") or 0) for item in project_rows),
        "subtask_count": sum(int(item.get("subtask_count") or 0) for item in project_rows),
    }

    options = {
        "projects": sorted({
            _text(item.get("project_key")).upper()
            for item in epics + stories + subtasks
            if _text(item.get("project_key"))
        }),
        "statuses": sorted({
            _text(item.get("status"))
            for item in epics + stories + subtasks
            if _text(item.get("status"))
        }),
        "assignees": sorted({
            _text(item.get("assignee"))
            for item in epics + stories + subtasks
            if _text(item.get("assignee"))
        }),
    }
    return final_rows, totals, options



def diff_snapshots(left: dict[str, Any], right: dict[str, Any]) -> dict[str, Any]:
    lt = left.get("totals", {}) if isinstance(left, dict) else {}
    rt = right.get("totals", {}) if isinstance(right, dict) else {}

    def _num(container: dict[str, Any], key: str) -> float:
        try:
            return float(container.get(key) or 0.0)
        except Exception:
            return 0.0

    return {
        "planned_hours_delta": round(_num(rt, "planned_hours") - _num(lt, "planned_hours"), 2),
        "actual_hours_delta": round(_num(rt, "actual_hours") - _num(lt, "actual_hours"), 2),
        "variance_hours_delta": round(_num(rt, "variance_hours") - _num(lt, "variance_hours"), 2),
        "project_count_delta": int(_num(rt, "project_count") - _num(lt, "project_count")),
        "epic_count_delta": int(_num(rt, "epic_count") - _num(lt, "epic_count")),
        "story_count_delta": int(_num(rt, "story_count") - _num(lt, "story_count")),
        "subtask_count_delta": int(_num(rt, "subtask_count") - _num(lt, "subtask_count")),
    }



def export_csv_bytes(snapshot: dict[str, Any]) -> bytes:
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow([
        "row_type",
        "project_key",
        "issue_key",
        "parent_key",
        "summary",
        "assignee",
        "status",
        "planned_start",
        "planned_due",
        "planned_hours",
        "actual_hours",
        "variance_hours",
    ])
    for row in snapshot.get("rows", []) or []:
        writer.writerow([
            _text(row.get("row_type")),
            _text(row.get("project_key")),
            _text(row.get("issue_key")),
            _text(row.get("parent_key")),
            _text(row.get("summary")),
            _text(row.get("assignee")),
            _text(row.get("status")),
            _text(row.get("planned_start")),
            _text(row.get("planned_due")),
            row.get("planned_hours", 0),
            row.get("actual_hours", 0),
            row.get("variance_hours", 0),
        ])
    return out.getvalue().encode("utf-8")



def export_xlsx_bytes(snapshot: dict[str, Any]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "planned_actual_table_view"
    ws.append([
        "row_type",
        "project_key",
        "issue_key",
        "parent_key",
        "summary",
        "assignee",
        "status",
        "planned_start",
        "planned_due",
        "planned_hours",
        "actual_hours",
        "variance_hours",
    ])
    for row in snapshot.get("rows", []) or []:
        ws.append([
            _text(row.get("row_type")),
            _text(row.get("project_key")),
            _text(row.get("issue_key")),
            _text(row.get("parent_key")),
            _text(row.get("summary")),
            _text(row.get("assignee")),
            _text(row.get("status")),
            _text(row.get("planned_start")),
            _text(row.get("planned_due")),
            float(row.get("planned_hours") or 0.0),
            float(row.get("actual_hours") or 0.0),
            float(row.get("variance_hours") or 0.0),
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
