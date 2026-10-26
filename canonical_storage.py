"""Replace obsolete canonical snapshots without touching user-maintained tables."""
from contextlib import closing
from pathlib import Path
import shutil
import sqlite3
import time
import json
import os
import subprocess
import sys


def _trim_cifs_allocation(db_path: Path) -> dict:
    """Release Azure SMB's retained allocation without replacing the database inode.

    Run in a child process: closing a raw descriptor in the server process can
    release POSIX SQLite locks belonging to its other connections.
    """
    if sys.platform != "linux":
        return {}
    completed = subprocess.run(
        [sys.executable, str(Path(__file__).resolve()), "--trim-cifs", str(db_path.resolve())],
        capture_output=True, text=True, check=True,
    )
    return json.loads(completed.stdout)


def _trim_cifs_in_process(db_path: Path) -> dict:
    import ctypes
    import errno
    import fcntl
    import hashlib
    import struct

    # Open before acquiring SQLite's lock; close only after releasing it.
    fd = os.open(db_path, os.O_RDWR)
    try:
        def allocation():
            buf = bytearray(struct.pack("<6I", 1, 5, 0, 0, 24, 0) + bytes(24))
            try:
                fcntl.ioctl(fd, 0xC018CF07, buf, True)  # CIFS_QUERY_INFO/FileStandardInformation
            except OSError as exc:
                if exc.errno in (errno.ENOTTY, errno.EOPNOTSUPP, errno.EINVAL):
                    return None
                raise
            return struct.unpack("<QQ", buf[24:40])

        with closing(sqlite3.connect(db_path, timeout=60)) as conn:
            conn.execute("BEGIN EXCLUSIVE")
            try:
                initial = allocation()
                if initial is None:
                    return {}
                allocated, size = initial
                result = {"allocated_bytes_before": allocated, "allocated_bytes_after": allocated}
                start = ((size + 1048575) // 1048576) * 1048576
                if allocated <= start:
                    return result

                def digest():
                    value = hashlib.sha256()
                    offset = 0
                    while offset < size:
                        block = os.pread(fd, min(1048576, size - offset), offset)
                        if not block:
                            raise OSError("Unexpected end of database during allocation verification")
                        value.update(block)
                        offset += len(block)
                    return value.digest()

                before = digest()
                libc = ctypes.CDLL(None, use_errno=True)
                libc.fallocate.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_longlong, ctypes.c_longlong]
                # Azure retains allocation after truncate, and ignores a hole
                # beyond EOF. Expose only the unused tail, punch it, restore EOF.
                # The valid prefix stays untouched and the inode stays in place.
                try:
                    os.ftruncate(fd, allocated)
                    if libc.fallocate(fd, 3, start, allocated - start) != 0:
                        code = ctypes.get_errno()
                        raise OSError(code, os.strerror(code))
                finally:
                    os.ftruncate(fd, size)
                    os.fsync(fd)
                if digest() != before:
                    raise OSError("Database bytes changed during allocation trimming")
                result["allocated_bytes_after"] = allocation()[0]
                if result["allocated_bytes_after"] > start:
                    raise OSError("Azure SMB still retains unused database allocation")
                return result
            finally:
                conn.rollback()
    finally:
        os.close(fd)


SNAPSHOT_TABLES = (
    "canonical_issue_scope_reasons", "canonical_worklogs", "canonical_issue_links",
    "canonical_issue_actuals", "canonical_assignee_period_hours",
    "canonical_issue_planning_flags", "canonical_hierarchy_summary",
    "canonical_project_assignee_summary", "canonical_issues",
)
DELETE_BATCH_ROWS = 5000


def _delete_in_batches(conn: sqlite3.Connection, table: str, key: str, current: str) -> int:
    """Bound journal growth; committed free pages can serve subsequent batches."""
    deleted = 0
    while True:
        try:
            with conn:
                count = conn.execute(
                    f"DELETE FROM {table} WHERE rowid IN "
                    f"(SELECT rowid FROM {table} WHERE {key} != ? LIMIT {DELETE_BATCH_ROWS})", (current,)
                ).rowcount
        except sqlite3.Error as exc:
            raise type(exc)(f"{table}: cleanup stopped after {deleted} committed rows: {exc}") from exc
        deleted += count
        if count < DELETE_BATCH_ROWS:
            return deleted
        # Let status/report readers acquire a shared lock between DELETE-mode
        # transactions on Azure's network share instead of starving them.
        time.sleep(0.01)


def replace_previous_snapshots(db_path: Path, run_id: str, compute_run_id: str) -> dict:
    """Called under the refresh lock, only after successful report promotion.

    Cleanup commits in bounded batches before VACUUM. Compaction failure is a
    maintenance warning, not a failed refresh; SQLite can reuse freed pages.
    The caller holds the refresh lock throughout validation and cleanup.
    """
    result = {"status": "pending", "bytes_before": db_path.stat().st_size, "deleted_rows": {}}
    with closing(sqlite3.connect(db_path, timeout=5)) as conn, conn:
        # Keep hot B-tree pages in memory on Azure's high-latency share. Dirty
        # page spilling can take an exclusive lock before COMMIT; keep it off
        # for this bounded maintenance transaction so readers can still run.
        conn.execute("PRAGMA cache_size=-65536")
        conn.execute("PRAGMA cache_spill=OFF")
        # The Azure build defaults to secure_delete=ON. FAST avoids rewriting
        # every discarded overflow page before VACUUM removes those pages.
        conn.execute("PRAGMA secure_delete=FAST")
        conn.execute("BEGIN IMMEDIATE")
        state = conn.execute(
            """SELECT last_success_run_id, active_run_id, active_fetch_run_id,
                      active_compute_run_id FROM canonical_refresh_state WHERE id=1"""
        ).fetchone()
        run = conn.execute("SELECT status FROM canonical_refresh_runs WHERE run_id=?", (run_id,)).fetchone()
        if not state or state[0] != run_id or not run or run[0] != "success":
            raise ValueError("Replacement requires the successfully published snapshot.")
        for table, key, current, active in (
            ("canonical_refresh_runs", "run_id", run_id, state[1]),
            ("canonical_fetch_runs", "fetch_run_id", run_id, state[2]),
            ("canonical_compute_runs", "compute_run_id", compute_run_id, state[3]),
        ):
            # Historical rows can retain 'running' after an interrupted request.
            # Only the lifecycle's active pointer identifies a live job. The
            # caller's refresh lock excludes another worker in this process.
            if conn.execute(
                f"SELECT 1 FROM {table} WHERE {key} = ? AND {key} != ? "
                "AND status IN ('running', 'cancel_requested') LIMIT 1",
                (active, current),
            ).fetchone():
                raise ValueError("Replacement deferred because another Fetch or Compute is running.")
        for table in SNAPSHOT_TABLES:
            result["deleted_rows"][table] = _delete_in_batches(conn, table, "run_id", run_id)
        for table, key, current in (
            ("canonical_compute_runs", "compute_run_id", compute_run_id),
            ("canonical_fetch_runs", "fetch_run_id", run_id),
            ("canonical_refresh_runs", "run_id", run_id),
            ("employee_performance_scoped_runs", "source_global_run_id", run_id),
        ):
            _delete_in_batches(conn, table, key, current)
        conn.execute(
            """UPDATE canonical_refresh_state SET last_success_fetch_run_id=?,
               last_success_compute_run_id=? WHERE id=1""", (run_id, compute_run_id)
        )
    result["status"] = "pruned"
    try:
        # Deleted pages still inflate the file until VACUUM succeeds. Estimate
        # working space from pages that remain in use, rather than rejecting
        # compaction solely because the old snapshots made the file enormous.
        # SQLite can still reject this estimate; that remains a visible warning.
        with closing(sqlite3.connect(db_path, timeout=5)) as conn:
            pages = conn.execute("PRAGMA page_count").fetchone()[0]
            free_pages = conn.execute("PRAGMA freelist_count").fetchone()[0]
            page_size = conn.execute("PRAGMA page_size").fetchone()[0]
        required = 2 * (pages - free_pages) * page_size
        result["compaction_estimated_bytes"] = required
        if shutil.disk_usage(db_path.parent).free < required:
            raise OSError(f"Compaction needs an estimated {required / (1024 ** 3):.1f} GiB free on the database drive.")
        with closing(sqlite3.connect(db_path, timeout=5)) as conn:
            conn.execute("PRAGMA temp_store=FILE")
            conn.execute("VACUUM")
        result.update(_trim_cifs_allocation(db_path))
        result["status"] = "compacted"
    except (OSError, sqlite3.Error, subprocess.SubprocessError, ValueError) as exc:
        result["warning"] = f"Old snapshots removed; disk compaction pending: {exc}"
    result["bytes_after"] = db_path.stat().st_size
    result["bytes_reclaimed"] = max(0,
        result.get("allocated_bytes_before", result["bytes_before"]) -
        result.get("allocated_bytes_after", result["bytes_after"]))
    return result


if __name__ == "__main__":
    if len(sys.argv) != 3 or sys.argv[1] != "--trim-cifs":
        raise SystemExit("Usage: canonical_storage.py --trim-cifs DATABASE")
    print(json.dumps(_trim_cifs_in_process(Path(sys.argv[2]))))
