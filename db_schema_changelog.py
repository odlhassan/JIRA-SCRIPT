"""
db_schema_changelog.py

Tracks every structural change (ADD COLUMN, DROP COLUMN, RENAME COLUMN, ADD TABLE,
DROP TABLE, RENAME TABLE) made to any SQLite database in this project.

Stored in:  db_schema_changelog.db  (project root, gitignored for production safety)

Key functions:
  init_changelog_db(path)              – Create/migrate the changelog DB itself.
  record_change(...)                   – Record one schema change event.
  snapshot_current_schema(db_path)     – Snapshot full schema of a target DB.
  get_all_changes(db_file)             – Return all change records for a DB file.
  get_pending_changes(db_file, since)  – Changes after a given schema_version.
  get_latest_snapshot(db_file)         – Most recent snapshot for a DB file.

Agents/developers must call record_change() whenever they ALTER or CREATE/DROP tables
in assignee_hours_capacity.db, epics_management.db, jira_exports.db, or any other
project SQLite database.

The changelog is the authoritative record used by db_migration.py to generate
migration scripts for production databases that are older than local.
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_CHANGELOG_DB = Path(__file__).resolve().parent / "db_schema_changelog.db"

# Operations tracked
OP_ADD_COLUMN = "ADD_COLUMN"
OP_DROP_COLUMN = "DROP_COLUMN"
OP_RENAME_COLUMN = "RENAME_COLUMN"
OP_MODIFY_COLUMN = "MODIFY_COLUMN"  # type or constraint change
OP_ADD_TABLE = "ADD_TABLE"
OP_DROP_TABLE = "DROP_TABLE"
OP_RENAME_TABLE = "RENAME_TABLE"
OP_ADD_INDEX = "ADD_INDEX"
OP_DROP_INDEX = "DROP_INDEX"


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _next_version(conn: sqlite3.Connection, db_file: str) -> int:
    row = conn.execute(
        "SELECT MAX(schema_version) FROM schema_change_log WHERE db_file = ?",
        (db_file,),
    ).fetchone()
    current = row[0] if row and row[0] is not None else 0
    return current + 1


# ---------------------------------------------------------------------------
# Initialise changelog DB
# ---------------------------------------------------------------------------

def init_changelog_db(changelog_path: Path = DEFAULT_CHANGELOG_DB) -> None:
    """Create or migrate the schema changelog database itself."""
    changelog_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(changelog_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS schema_change_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                change_id TEXT NOT NULL UNIQUE,
                db_file TEXT NOT NULL,
                schema_version INTEGER NOT NULL,
                operation TEXT NOT NULL,
                table_name TEXT NOT NULL DEFAULT '',
                column_name TEXT NOT NULL DEFAULT '',
                old_column_name TEXT NOT NULL DEFAULT '',
                data_type TEXT NOT NULL DEFAULT '',
                nullable INTEGER NOT NULL DEFAULT 1,
                default_value TEXT NOT NULL DEFAULT '',
                reason TEXT NOT NULL DEFAULT '',
                files_reading_json TEXT NOT NULL DEFAULT '[]',
                referencing_tables_json TEXT NOT NULL DEFAULT '[]',
                previous_state_json TEXT NOT NULL DEFAULT '{}',
                new_state_json TEXT NOT NULL DEFAULT '{}',
                full_ddl TEXT NOT NULL DEFAULT '',
                changed_at_utc TEXT NOT NULL,
                changed_by TEXT NOT NULL DEFAULT 'agent',
                migration_applied_at_utc TEXT NOT NULL DEFAULT '',
                notes TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS schema_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                db_file TEXT NOT NULL,
                schema_version INTEGER NOT NULL,
                snapshot_json TEXT NOT NULL,
                created_at_utc TEXT NOT NULL,
                UNIQUE(db_file, schema_version)
            );

            CREATE TABLE IF NOT EXISTS migration_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id TEXT NOT NULL UNIQUE,
                target_db_file TEXT NOT NULL,
                from_version INTEGER NOT NULL DEFAULT 0,
                to_version INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'running',
                steps_total INTEGER NOT NULL DEFAULT 0,
                steps_done INTEGER NOT NULL DEFAULT 0,
                log_json TEXT NOT NULL DEFAULT '[]',
                error_text TEXT NOT NULL DEFAULT '',
                started_at_utc TEXT NOT NULL,
                finished_at_utc TEXT NOT NULL DEFAULT ''
            );
            """
        )
        conn.commit()


# ---------------------------------------------------------------------------
# Record a schema change
# ---------------------------------------------------------------------------

def record_change(
    db_file: str,
    operation: str,
    table_name: str,
    *,
    column_name: str = "",
    old_column_name: str = "",
    data_type: str = "",
    nullable: bool = True,
    default_value: str = "",
    reason: str = "",
    files_reading: list[str] | None = None,
    referencing_tables: list[str] | None = None,
    previous_state: dict[str, Any] | None = None,
    new_state: dict[str, Any] | None = None,
    full_ddl: str = "",
    changed_by: str = "agent",
    notes: str = "",
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> str:
    """Record a schema change.  Returns the new change_id (UUID)."""
    init_changelog_db(changelog_path)
    change_id = str(uuid.uuid4())
    now = _utc_now()
    with sqlite3.connect(changelog_path) as conn:
        version = _next_version(conn, db_file)
        conn.execute(
            """
            INSERT INTO schema_change_log (
                change_id, db_file, schema_version, operation,
                table_name, column_name, old_column_name,
                data_type, nullable, default_value,
                reason, files_reading_json, referencing_tables_json,
                previous_state_json, new_state_json, full_ddl,
                changed_at_utc, changed_by, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                change_id, db_file, version, operation,
                table_name, column_name, old_column_name,
                data_type, 1 if nullable else 0, default_value,
                reason,
                json.dumps(files_reading or []),
                json.dumps(referencing_tables or []),
                json.dumps(previous_state or {}),
                json.dumps(new_state or {}),
                full_ddl,
                now, changed_by, notes,
            ),
        )
        conn.commit()
    return change_id


# ---------------------------------------------------------------------------
# Snapshot the schema of a target DB
# ---------------------------------------------------------------------------

def snapshot_current_schema(
    db_path: Path,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> int:
    """
    Introspect all tables/columns in db_path, store as JSON snapshot.
    Returns the schema_version assigned.
    """
    init_changelog_db(changelog_path)
    db_file = db_path.name

    schema: dict[str, Any] = {"tables": {}, "indexes": []}
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        )
        tables = cur.fetchall()
        for (tname, tsql) in tables:
            cur.execute(f"PRAGMA table_info({tname})")
            cols = [
                {
                    "cid": r[0], "name": r[1], "type": r[2],
                    "notnull": r[3], "dflt_value": r[4], "pk": r[5],
                }
                for r in cur.fetchall()
            ]
            cur.execute(f"PRAGMA foreign_key_list({tname})")
            fks = [
                {"id": r[0], "seq": r[1], "table": r[2], "from": r[3], "to": r[4]}
                for r in cur.fetchall()
            ]
            schema["tables"][tname] = {"columns": cols, "foreign_keys": fks, "ddl": tsql or ""}

        cur.execute(
            "SELECT name, tbl_name, sql FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        )
        schema["indexes"] = [
            {"name": r[0], "table": r[1], "ddl": r[2] or ""}
            for r in cur.fetchall()
        ]

    now = _utc_now()
    with sqlite3.connect(changelog_path) as conn:
        version = _next_version(conn, db_file)
        conn.execute(
            """
            INSERT OR REPLACE INTO schema_snapshots (db_file, schema_version, snapshot_json, created_at_utc)
            VALUES (?, ?, ?, ?)
            """,
            (db_file, version, json.dumps(schema), now),
        )
        conn.commit()

    return version


# ---------------------------------------------------------------------------
# Query helpers
# ---------------------------------------------------------------------------

def get_all_changes(
    db_file: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> list[dict[str, Any]]:
    """Return all change log rows for a given db_file, oldest first."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM schema_change_log WHERE db_file = ? ORDER BY schema_version ASC",
            (db_file,),
        ).fetchall()
    return [dict(r) for r in rows]


def get_pending_changes(
    db_file: str,
    since_version: int = 0,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> list[dict[str, Any]]:
    """Return changes with schema_version > since_version, oldest first."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT * FROM schema_change_log
            WHERE db_file = ? AND schema_version > ? AND migration_applied_at_utc = ''
            ORDER BY schema_version ASC
            """,
            (db_file, since_version),
        ).fetchall()
    return [dict(r) for r in rows]


def get_latest_snapshot(
    db_file: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> dict[str, Any] | None:
    """Return the most recent schema snapshot for a db_file."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            """
            SELECT * FROM schema_snapshots WHERE db_file = ?
            ORDER BY schema_version DESC LIMIT 1
            """,
            (db_file,),
        ).fetchone()
    if row is None:
        return None
    result = dict(row)
    result["schema"] = json.loads(result.get("snapshot_json") or "{}")
    return result


def get_schema_version(
    db_file: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> int:
    """Return the current (latest) schema_version for db_file. 0 if none."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        row = conn.execute(
            "SELECT MAX(schema_version) FROM schema_change_log WHERE db_file = ?",
            (db_file,),
        ).fetchone()
    return row[0] if row and row[0] is not None else 0


def list_tracked_databases(
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> list[dict[str, Any]]:
    """Return summary of all tracked databases."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        rows = conn.execute(
            """
            SELECT db_file,
                   MAX(schema_version) AS current_version,
                   COUNT(*) AS change_count,
                   MAX(changed_at_utc) AS last_changed_at
            FROM schema_change_log
            GROUP BY db_file
            ORDER BY db_file
            """,
        ).fetchall()
    return [
        {"db_file": r[0], "current_version": r[1], "change_count": r[2], "last_changed_at": r[3]}
        for r in rows
    ]


def mark_change_applied(
    change_id: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> None:
    """Mark a change record as applied (migration completed)."""
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.execute(
            "UPDATE schema_change_log SET migration_applied_at_utc = ? WHERE change_id = ?",
            (_utc_now(), change_id),
        )
        conn.commit()


def insert_migration_run(
    run_id: str,
    target_db_file: str,
    from_version: int,
    to_version: int,
    steps_total: int,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> None:
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.execute(
            """
            INSERT INTO migration_runs
            (run_id, target_db_file, from_version, to_version, status, steps_total, started_at_utc)
            VALUES (?, ?, ?, ?, 'running', ?, ?)
            """,
            (run_id, target_db_file, from_version, to_version, steps_total, _utc_now()),
        )
        conn.commit()


def update_migration_run(
    run_id: str,
    *,
    steps_done: int | None = None,
    status: str | None = None,
    log_entry: dict[str, Any] | None = None,
    error_text: str | None = None,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> None:
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        row = conn.execute(
            "SELECT log_json, steps_done FROM migration_runs WHERE run_id = ?", (run_id,)
        ).fetchone()
        if not row:
            return
        existing_log = json.loads(row[0] or "[]")
        current_done = row[1]

        if log_entry:
            existing_log.append(log_entry)
        updates = []
        params: list[Any] = []
        if steps_done is not None:
            updates.append("steps_done = ?")
            params.append(steps_done)
        if status is not None:
            updates.append("status = ?")
            params.append(status)
            if status in ("success", "error"):
                updates.append("finished_at_utc = ?")
                params.append(_utc_now())
        if error_text is not None:
            updates.append("error_text = ?")
            params.append(error_text)
        updates.append("log_json = ?")
        params.append(json.dumps(existing_log))
        params.append(run_id)
        conn.execute(
            f"UPDATE migration_runs SET {', '.join(updates)} WHERE run_id = ?",
            params,
        )
        conn.commit()


def get_migration_run(
    run_id: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> dict[str, Any] | None:
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM migration_runs WHERE run_id = ?", (run_id,)
        ).fetchone()
    if row is None:
        return None
    result = dict(row)
    result["log"] = json.loads(result.get("log_json") or "[]")
    return result


def get_latest_migration_run(
    target_db_file: str,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> dict[str, Any] | None:
    init_changelog_db(changelog_path)
    with sqlite3.connect(changelog_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            """
            SELECT * FROM migration_runs WHERE target_db_file = ?
            ORDER BY id DESC LIMIT 1
            """,
            (target_db_file,),
        ).fetchone()
    if row is None:
        return None
    result = dict(row)
    result["log"] = json.loads(result.get("log_json") or "[]")
    return result
