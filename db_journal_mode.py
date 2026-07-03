"""Filesystem-aware SQLite journal mode selection.

WAL mode has been fully disabled in this project. Azure App Service mounts the
app's writable `/home` directory over an SMB/CIFS network share, where SQLite's
WAL mode (shared-memory index backed by mmap in a `-shm` file) is unreliable and
produced "database disk image is malformed" corruption on production canonical
refreshes (see the `recovery/` incident this fixed). Rather than keep two code
paths (WAL locally, rollback journal on Azure), every environment now uses the
rollback journal (`DELETE` mode) so local behavior always matches production and
no `-wal`/`-shm` sidecar files are ever created.

Override via environment (kept for diagnostics/testing only):
- ``EPR_FORCE_WAL=1``   -> force WAL anywhere (not recommended; for debugging only).
- ``EPR_DISABLE_WAL=1`` -> explicit no-op; DELETE mode is already always used.
"""

from __future__ import annotations

import os
import sqlite3


def _is_truthy(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "yes", "y", "on"}


def is_network_filesystem_host() -> bool:
    """Retained for compatibility; WAL is disabled everywhere regardless of host."""
    if _is_truthy(os.getenv("EPR_FORCE_WAL")):
        return False
    return True


def safe_journal_mode() -> str:
    if _is_truthy(os.getenv("EPR_FORCE_WAL")):
        return "WAL"
    return "DELETE"


def apply_journal_mode(conn: sqlite3.Connection) -> str:
    """Set the journal mode (always DELETE unless EPR_FORCE_WAL=1) and return it."""
    mode = safe_journal_mode()
    try:
        conn.execute(f"PRAGMA journal_mode={mode}")
    except sqlite3.OperationalError as exc:
        message = str(exc).lower()
        if "database is locked" in message:
            return mode
        raise
    return mode
