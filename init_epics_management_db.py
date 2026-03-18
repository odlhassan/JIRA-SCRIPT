"""
Initialize (or migrate) the Epics Management database used by Epics Planner.

Creates assignee_hours_capacity.db and the epics_management tables if missing,
and applies schema migrations (e.g. is_sealed, epics_management_approved_dates).

Usage:
  python init_epics_management_db.py
  python init_epics_management_db.py --db path/to/assignee_hours_capacity.db

Env:
  JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH  Override DB path (default: assignee_hours_capacity.db in project root).
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

# Ensure project root is on path so report_server can be imported
_project_root = Path(__file__).resolve().parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from report_server import _init_epics_management_db, _resolve_capacity_runtime_paths


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Initialize or migrate the Epics Management database (Epics Planner)."
    )
    parser.add_argument(
        "--db",
        default=os.getenv("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", "").strip(),
        help="Path to assignee_hours_capacity.db. Default: env JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH or assignee_hours_capacity.db in project root.",
    )
    args = parser.parse_args()

    base_dir = _project_root
    if args.db:
        db_path = Path(args.db)
        if not db_path.is_absolute():
            db_path = base_dir / db_path
    else:
        paths = _resolve_capacity_runtime_paths(base_dir)
        db_path = paths["db_path"]

    _init_epics_management_db(db_path)
    print(f"Epics management DB initialized: {db_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
