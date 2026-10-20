"""Replace obsolete canonical snapshots without touching user-maintained tables."""
from contextlib import closing
from pathlib import Path
import shutil
import sqlite3


SNAPSHOT_TABLES = (
    "canonical_issue_scope_reasons", "canonical_worklogs", "canonical_issue_links",
    "canonical_issue_actuals", "canonical_assignee_period_hours",
    "canonical_issue_planning_flags", "canonical_hierarchy_summary",
    "canonical_project_assignee_summary", "canonical_issues",
)


def replace_previous_snapshots(db_path: Path, run_id: str, compute_run_id: str) -> dict:
    """Called under the refresh lock, only after successful report promotion.

    Cleanup commits before VACUUM. Compaction failure is a maintenance warning,
    not a failed refresh; SQLite can reuse the freed pages in the meantime.
    """
    result = {"status": "pending", "bytes_before": db_path.stat().st_size, "deleted_rows": {}}
    with closing(sqlite3.connect(db_path, timeout=5)) as conn, conn:
        conn.execute("BEGIN IMMEDIATE")
        state = conn.execute(
            "SELECT last_success_run_id FROM canonical_refresh_state WHERE id=1"
        ).fetchone()
        run = conn.execute("SELECT status FROM canonical_refresh_runs WHERE run_id=?", (run_id,)).fetchone()
        if not state or state[0] != run_id or not run or run[0] != "success":
            raise ValueError("Replacement requires the successfully published snapshot.")
        for table, key, current in (
            ("canonical_refresh_runs", "run_id", run_id),
            ("canonical_fetch_runs", "fetch_run_id", run_id),
            ("canonical_compute_runs", "compute_run_id", compute_run_id),
        ):
            if conn.execute(
                f"SELECT 1 FROM {table} WHERE {key} != ? AND status IN ('running', 'cancel_requested') LIMIT 1",
                (current,),
            ).fetchone():
                raise ValueError("Replacement deferred because another Fetch or Compute is running.")
        for table in SNAPSHOT_TABLES:
            result["deleted_rows"][table] = conn.execute(
                f"DELETE FROM {table} WHERE run_id != ?", (run_id,)
            ).rowcount
        conn.execute("DELETE FROM canonical_compute_runs WHERE compute_run_id != ?", (compute_run_id,))
        conn.execute("DELETE FROM canonical_fetch_runs WHERE fetch_run_id != ?", (run_id,))
        conn.execute("DELETE FROM canonical_refresh_runs WHERE run_id != ?", (run_id,))
        conn.execute("DELETE FROM employee_performance_scoped_runs WHERE source_global_run_id != ?", (run_id,))
        conn.execute(
            """UPDATE canonical_refresh_state SET last_success_fetch_run_id=?,
               last_success_compute_run_id=? WHERE id=1""", (run_id, compute_run_id)
        )
    result["status"] = "pruned"
    try:
        # SQLite documents up to twice the original DB size in temporary space.
        required = 2 * db_path.stat().st_size
        if shutil.disk_usage(db_path.parent).free < required:
            raise OSError(f"Compaction needs up to {required / (1024 ** 3):.1f} GiB free on the database drive.")
        with closing(sqlite3.connect(db_path, timeout=5)) as conn:
            conn.execute("PRAGMA temp_store=FILE")
            conn.execute("VACUUM")
        result["status"] = "compacted"
    except (OSError, sqlite3.Error) as exc:
        result["warning"] = f"Old snapshots removed; disk compaction pending: {exc}"
    result["bytes_after"] = db_path.stat().st_size
    result["bytes_reclaimed"] = max(0, result["bytes_before"] - result["bytes_after"])
    return result
