"""Idempotent migration: ensure the standalone project_images database exists.

This feature stores per-project thumbnail + logo metadata in a dedicated
``project_images.db`` (with the image bytes as files alongside it) instead of
adding columns to the multi-gigabyte ``assignee_hours_capacity.db``. That keeps
production upgrades cheap: there is nothing to download/rewrite on the large DB.

The application already calls ``init_project_images_db`` at startup, so on Azure
the DB is created automatically on first boot after deploy. This script exists so
the migration path is explicit and can be run manually against the persistent
image directory if desired. It is fully idempotent (CREATE TABLE IF NOT EXISTS).

Usage (from repo root):
    python migrations/2026-07-23_project_images.py
    python migrations/2026-07-23_project_images.py --dir /home/data/project_images
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from project_image_registry import (  # noqa: E402
    init_project_images_db,
    resolve_project_image_paths,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Ensure project_images.db exists (idempotent).")
    parser.add_argument(
        "--dir",
        default="",
        help="Image directory to initialize. Defaults to the app-resolved location "
        "(JIRA_PROJECT_IMAGE_DIR, else <base>/data/project_images, else $HOME/data/project_images).",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent.parent
    if args.dir:
        images_dir = Path(args.dir)
        images_dir.mkdir(parents=True, exist_ok=True)
        db_path = images_dir / "project_images.db"
    else:
        resolved = resolve_project_image_paths(base_dir)
        images_dir = resolved["images_dir"]
        db_path = resolved["db_path"]

    init_project_images_db(db_path)
    print(f"OK: project_images.db ready at {db_path}")
    print(f"    image files directory: {images_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
