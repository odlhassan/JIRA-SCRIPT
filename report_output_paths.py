"""Shared resolution of the directory that generated report artifacts are written to.

Every generator in this repo anchors its inputs and outputs to its own source
directory (``Path(__file__).resolve().parent``). That is the repo root locally, but
on Azure App Service ``WEBSITE_RUN_FROM_PACKAGE`` mounts ``/home/site/wwwroot``
read-only, so writing ``rlt_leave_report.xlsx`` / ``rlt_leave_report.html`` there
fails with ``OSError: [Errno 30] Read-only file system`` and takes the whole
Colossal Refresh Compute phase down with it.

``resolve_output_base()`` mirrors ``report_server._canonical_bridge_artifact_base_dir()``
exactly, so generators read and write the same artifact set the canonical
compatibility bridge produces. On a writable root it returns the script directory
unchanged, which keeps local behaviour identical.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

ARTIFACT_DIR_ENV = "JIRA_CANONICAL_ARTIFACT_DIR"


def is_writable_directory(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write-probe"
        with open(probe, "a", encoding="utf-8"):
            pass
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def resolve_output_base(script_dir: Path) -> Path:
    script_dir = Path(script_dir)
    configured = (os.getenv(ARTIFACT_DIR_ENV) or "").strip()
    if configured:
        path = Path(configured)
        if not path.is_absolute():
            path = script_dir / path
    elif is_writable_directory(script_dir):
        return script_dir
    else:
        path = Path(os.getenv("HOME") or tempfile.gettempdir()) / "data" / "canonical_artifacts"
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError:
        return script_dir
    return path
