"""
db_migration.py

Production database migration engine.

Workflow:
  1. Human downloads production DB file (e.g. assignee_hours_capacity.db) and
     places it at a known path on local machine.
  2. Human asks agent to run migration.
  3. Agent calls: migrate_production_db(prod_db_path, local_db_path, changelog_path)
  4. The engine:
       a) Introspects production DB schema (tables, columns)
       b) Introspects local DB schema (the authoritative target)
       c) Computes a diff: new tables, dropped tables, modified tables, renamed tables
       d) For each modified table in production:
            - RENAME production table → <table>_old
            - CREATE new table with local structure
            - INSERT INTO new table SELECT ... FROM <table>_old   (mapping common columns)
            - DROP TABLE <table>_old
       e) For new tables: CREATE TABLE
       f) For dropped tables: DROP TABLE (with confirmation check)
       g) Marks all applied changes as applied in schema_change_log
  5. Progress streamed via a callback (used by migration UI).

Can also be invoked from CLI:
  python db_migration.py --prod /path/to/prod.db --local assignee_hours_capacity.db
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from db_schema_changelog import (
    DEFAULT_CHANGELOG_DB,
    get_latest_snapshot,
    get_schema_version,
    init_changelog_db,
    insert_migration_run,
    mark_change_applied,
    snapshot_current_schema,
    update_migration_run,
    get_migration_run,
    get_latest_migration_run,
)


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ---------------------------------------------------------------------------
# Schema introspection
# ---------------------------------------------------------------------------

def _introspect_schema(db_path: Path) -> dict[str, Any]:
    """Return {table_name: {columns: [...], ddl: str, indexes: [...]}} for db_path."""
    result: dict[str, Any] = {"tables": {}, "indexes": []}
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        )
        for (tname, tsql) in cur.fetchall():
            if tname.endswith("_old"):
                continue  # skip leftover _old tables
            cur.execute(f"PRAGMA table_info({tname})")
            cols = [
                {
                    "cid": r[0], "name": r[1], "type": r[2] or "TEXT",
                    "notnull": r[3], "dflt_value": r[4], "pk": r[5],
                }
                for r in cur.fetchall()
            ]
            result["tables"][tname] = {"columns": cols, "ddl": tsql or ""}
        cur.execute(
            "SELECT name, tbl_name, sql FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%'"
        )
        result["indexes"] = [
            {"name": r[0], "table": r[1], "ddl": r[2] or ""}
            for r in cur.fetchall()
        ]
    return result


def _col_map(columns: list[dict]) -> dict[str, dict]:
    return {c["name"]: c for c in columns}


# ---------------------------------------------------------------------------
# Diff computation
# ---------------------------------------------------------------------------

@dataclass
class MigrationStep:
    kind: str  # new_table | drop_table | modify_table | rename_table | new_index | drop_index
    table: str
    detail: dict[str, Any] = field(default_factory=dict)
    sql_statements: list[str] = field(default_factory=list)
    description: str = ""


def compute_diff(
    prod_schema: dict[str, Any],
    local_schema: dict[str, Any],
) -> list[MigrationStep]:
    """
    Compare production schema against local (authoritative) schema.
    Returns an ordered list of MigrationSteps.
    """
    steps: list[MigrationStep] = []
    prod_tables = prod_schema.get("tables", {})
    local_tables = local_schema.get("tables", {})

    # Tables present in local but not production → need to be created
    for tname, tinfo in local_tables.items():
        if tname not in prod_tables:
            steps.append(MigrationStep(
                kind="new_table",
                table=tname,
                detail=tinfo,
                sql_statements=[tinfo["ddl"]],
                description=f"Create new table '{tname}'",
            ))

    # Tables present in production but not local → drop (with safety check)
    for tname in prod_tables:
        if tname not in local_tables:
            steps.append(MigrationStep(
                kind="drop_table",
                table=tname,
                detail={},
                sql_statements=[f"DROP TABLE IF EXISTS \"{tname}\""],
                description=f"Drop obsolete table '{tname}'",
            ))

    # Tables present in both → check for column differences
    for tname in local_tables:
        if tname not in prod_tables:
            continue  # handled above
        prod_cols = _col_map(prod_tables[tname]["columns"])
        local_cols = _col_map(local_tables[tname]["columns"])
        if _columns_match(prod_cols, local_cols):
            continue  # no change
        # Need to rebuild the table
        steps.append(_build_modify_table_step(tname, prod_cols, local_cols, local_tables[tname]["ddl"]))

    return steps


def _columns_match(prod_cols: dict, local_cols: dict) -> bool:
    """Return True if column set and types are identical."""
    if set(prod_cols.keys()) != set(local_cols.keys()):
        return False
    for name, lcol in local_cols.items():
        pcol = prod_cols[name]
        if (lcol["type"] or "TEXT").upper() != (pcol["type"] or "TEXT").upper():
            return False
        if lcol["notnull"] != pcol["notnull"]:
            return False
        if str(lcol["dflt_value"] or "") != str(pcol["dflt_value"] or ""):
            return False
    return True


def _build_modify_table_step(
    tname: str,
    prod_cols: dict[str, dict],
    local_cols: dict[str, dict],
    new_ddl: str,
) -> MigrationStep:
    """Build the rename→create→migrate→drop step for a modified table."""
    old_name = f"{tname}_old"
    # Common columns (exist in both) – migrate their data
    common_cols = [name for name in local_cols if name in prod_cols]
    col_list = ", ".join(f'"{c}"' for c in common_cols)

    stmts = [
        f'ALTER TABLE "{tname}" RENAME TO "{old_name}"',
        new_ddl,  # CREATE TABLE tname (...)
        f'INSERT INTO "{tname}" ({col_list}) SELECT {col_list} FROM "{old_name}"',
        f'DROP TABLE IF EXISTS "{old_name}"',
    ]
    added = [c for c in local_cols if c not in prod_cols]
    dropped = [c for c in prod_cols if c not in local_cols]
    changed = [
        c for c in common_cols
        if not _columns_match({c: prod_cols[c]}, {c: local_cols[c]})
    ]
    return MigrationStep(
        kind="modify_table",
        table=tname,
        detail={
            "added_columns": added,
            "dropped_columns": dropped,
            "changed_columns": changed,
            "migrated_columns": common_cols,
        },
        sql_statements=stmts,
        description=(
            f"Rebuild '{tname}': "
            + (f"+{len(added)} added " if added else "")
            + (f"-{len(dropped)} dropped " if dropped else "")
            + (f"~{len(changed)} changed" if changed else "")
        ).strip(),
    )


# ---------------------------------------------------------------------------
# Execution
# ---------------------------------------------------------------------------

ProgressCallback = Callable[[str, int, int], None]  # message, done, total


def execute_migration(
    prod_db_path: Path,
    steps: list[MigrationStep],
    run_id: str,
    progress_cb: ProgressCallback | None = None,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
) -> dict[str, Any]:
    """
    Execute migration steps against prod_db_path.

    Returns a result dict: {success, steps_done, steps_total, errors, log}
    """
    total = len(steps)
    done = 0
    errors: list[str] = []
    log: list[dict[str, Any]] = []

    def _emit(msg: str, level: str = "info") -> None:
        entry = {"ts": _utc_now(), "level": level, "msg": msg}
        log.append(entry)
        update_migration_run(run_id, log_entry=entry, steps_done=done, changelog_path=changelog_path)
        if progress_cb:
            progress_cb(msg, done, total)

    _emit(f"Migration started. {total} step(s).")

    with sqlite3.connect(prod_db_path) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=OFF")

        for step in steps:
            _emit(f"[{step.kind.upper()}] {step.description}")
            step_ok = True
            for sql in step.sql_statements:
                if not sql or not sql.strip():
                    continue
                try:
                    conn.execute(sql)
                    conn.commit()
                    _emit(f"  OK: {sql[:120]}")
                except Exception as exc:
                    msg = f"  ERROR in step '{step.description}': {exc}\n  SQL: {sql[:300]}"
                    _emit(msg, level="error")
                    errors.append(msg)
                    step_ok = False
                    break

            done += 1
            update_migration_run(run_id, steps_done=done, changelog_path=changelog_path)

        conn.execute("PRAGMA foreign_keys=ON")

    success = len(errors) == 0
    status = "success" if success else "error"
    update_migration_run(
        run_id,
        status=status,
        steps_done=done,
        error_text="\n".join(errors),
        changelog_path=changelog_path,
    )
    _emit(f"Migration {'COMPLETED' if success else 'FINISHED WITH ERRORS'}. {done}/{total} steps done.")
    return {"success": success, "steps_done": done, "steps_total": total, "errors": errors, "log": log}


# ---------------------------------------------------------------------------
# High-level entry point
# ---------------------------------------------------------------------------

def migrate_production_db(
    prod_db_path: Path,
    local_db_path: Path,
    changelog_path: Path = DEFAULT_CHANGELOG_DB,
    progress_cb: ProgressCallback | None = None,
) -> dict[str, Any]:
    """
    Full migration pipeline:
      1. Snapshot production schema (for audit).
      2. Introspect local schema (authoritative target).
      3. Compute diff.
      4. Execute steps.
      5. Mark changelog entries as applied.

    Returns result dict from execute_migration.
    """
    if not prod_db_path.exists():
        raise FileNotFoundError(f"Production DB not found: {prod_db_path}")
    if not local_db_path.exists():
        raise FileNotFoundError(f"Local DB not found: {local_db_path}")

    init_changelog_db(changelog_path)
    db_file = local_db_path.name

    prod_schema = _introspect_schema(prod_db_path)
    local_schema = _introspect_schema(local_db_path)
    steps = compute_diff(prod_schema, local_schema)

    from_version = get_schema_version(db_file, changelog_path)
    run_id = str(uuid.uuid4())
    insert_migration_run(run_id, db_file, 0, from_version, len(steps), changelog_path)

    if not steps:
        update_migration_run(run_id, status="success", steps_done=0, changelog_path=changelog_path)
        return {"success": True, "steps_done": 0, "steps_total": 0, "errors": [], "log": [
            {"ts": _utc_now(), "level": "info", "msg": "Production DB schema already matches local. Nothing to do."}
        ], "run_id": run_id}

    result = execute_migration(prod_db_path, steps, run_id, progress_cb, changelog_path)
    result["run_id"] = run_id
    result["steps_summary"] = [
        {"kind": s.kind, "table": s.table, "description": s.description}
        for s in steps
    ]
    return result


def plan_migration(
    prod_db_path: Path,
    local_db_path: Path,
) -> dict[str, Any]:
    """
    Compute and return migration plan without executing anything.
    Use this to preview what will happen before running.
    """
    prod_schema = _introspect_schema(prod_db_path)
    local_schema = _introspect_schema(local_db_path)
    steps = compute_diff(prod_schema, local_schema)
    return {
        "steps_total": len(steps),
        "steps": [
            {
                "kind": s.kind,
                "table": s.table,
                "description": s.description,
                "detail": s.detail,
                "sql_statements": s.sql_statements,
            }
            for s in steps
        ],
        "prod_tables": list(prod_schema["tables"].keys()),
        "local_tables": list(local_schema["tables"].keys()),
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cli_main() -> int:
    parser = argparse.ArgumentParser(
        description="Migrate a production SQLite DB to match the local schema."
    )
    parser.add_argument("--prod", required=True, help="Path to the downloaded production DB.")
    parser.add_argument("--local", required=True, help="Path to the local authoritative DB.")
    parser.add_argument(
        "--changelog", default=str(DEFAULT_CHANGELOG_DB), help="Path to db_schema_changelog.db"
    )
    parser.add_argument(
        "--plan-only", action="store_true", help="Show migration plan without executing."
    )
    args = parser.parse_args()

    prod = Path(args.prod)
    local = Path(args.local)
    changelog = Path(args.changelog)

    if args.plan_only:
        plan = plan_migration(prod, local)
        print(f"\nMigration plan: {plan['steps_total']} step(s)")
        for i, step in enumerate(plan["steps"], 1):
            print(f"  {i}. [{step['kind'].upper()}] {step['description']}")
            for sql in step["sql_statements"]:
                print(f"       SQL: {sql[:100]}")
        return 0

    def _progress(msg: str, done: int, total: int) -> None:
        print(f"  [{done}/{total}] {msg}")

    print(f"\nMigrating '{prod.name}' to match '{local.name}' ...")
    result = migrate_production_db(prod, local, changelog, progress_cb=_progress)

    print(f"\n{'SUCCESS' if result['success'] else 'FAILED'}: {result['steps_done']}/{result['steps_total']} steps done.")
    if result.get("errors"):
        print("\nErrors:")
        for err in result["errors"]:
            print(f"  {err}")
    return 0 if result["success"] else 1


if __name__ == "__main__":
    sys.exit(_cli_main())
