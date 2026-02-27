from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


_PROJECT_KEY_PATTERN = re.compile(r"^[A-Z0-9_-]+$")
_COLOR_HEX_PATTERN = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_project_key(value: object) -> str:
    key = _to_text(value).upper()
    if not key:
        raise ValueError("project_key is required.")
    if not _PROJECT_KEY_PATTERN.match(key):
        raise ValueError("project_key must match ^[A-Z0-9_\\-]+$.")
    return key


def normalize_color_hex(value: object) -> str:
    text = _to_text(value)
    if not text:
        raise ValueError("color_hex is required.")
    if not _COLOR_HEX_PATTERN.match(text):
        raise ValueError("color_hex must be in #RRGGBB format.")
    return text.upper()


def normalize_managed_project_payload(payload: dict, require_all_fields: bool = True) -> dict:
    raw = payload or {}
    project_key = normalize_project_key(raw.get("project_key"))
    project_name = _to_text(raw.get("project_name"))
    display_name = _to_text(raw.get("display_name"))
    color_hex = normalize_color_hex(raw.get("color_hex"))
    is_active = 1 if bool(raw.get("is_active", True)) else 0

    if require_all_fields:
        if not project_name:
            raise ValueError("project_name is required.")
        if not display_name:
            raise ValueError("display_name is required.")
    else:
        if not project_name:
            project_name = project_key
        if not display_name:
            display_name = project_name

    return {
        "project_key": project_key,
        "project_name": project_name,
        "display_name": display_name,
        "color_hex": color_hex,
        "is_active": is_active,
    }


def deterministic_color_for_project_key(project_key: str) -> str:
    normalized = normalize_project_key(project_key)
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    # Keep colors visible and avoid extremely dark/light shades.
    r = 64 + (int(digest[0:2], 16) % 128)
    g = 64 + (int(digest[2:4], 16) % 128)
    b = 64 + (int(digest[4:6], 16) % 128)
    return f"#{r:02X}{g:02X}{b:02X}"


def parse_project_keys_from_env(raw: str | None = None) -> list[str]:
    text = raw
    if text is None:
        text = os.getenv("JIRA_PROJECT_KEYS", "")
    out: list[str] = []
    seen: set[str] = set()
    for item in _to_text(text).split(","):
        candidate = _to_text(item).upper()
        if not candidate:
            continue
        if not _PROJECT_KEY_PATTERN.match(candidate):
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        out.append(candidate)
    return out


def init_managed_projects_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS managed_projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_key TEXT NOT NULL UNIQUE,
                project_name TEXT NOT NULL,
                display_name TEXT NOT NULL,
                color_hex TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _row_to_project(row: sqlite3.Row) -> dict:
    return {
        "id": int(row["id"]),
        "project_key": _to_text(row["project_key"]),
        "project_name": _to_text(row["project_name"]),
        "display_name": _to_text(row["display_name"]),
        "color_hex": _to_text(row["color_hex"]).upper(),
        "is_active": bool(int(row["is_active"] or 0)),
        "created_at_utc": _to_text(row["created_at_utc"]),
        "updated_at_utc": _to_text(row["updated_at_utc"]),
    }


def _load_managed_project_by_key(db_path: Path, project_key: str) -> dict | None:
    init_managed_projects_db(db_path)
    key = normalize_project_key(project_key)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT id, project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
            FROM managed_projects
            WHERE project_key = ?
            """,
            (key,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    return _row_to_project(row)


def list_managed_projects(db_path: Path, include_inactive: bool = False) -> list[dict]:
    init_managed_projects_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        if include_inactive:
            rows = conn.execute(
                """
                SELECT id, project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
                FROM managed_projects
                ORDER BY lower(display_name) ASC, project_key ASC
                """
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT id, project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
                FROM managed_projects
                WHERE is_active = 1
                ORDER BY lower(display_name) ASC, project_key ASC
                """
            ).fetchall()
    finally:
        conn.close()
    return [_row_to_project(row) for row in rows]


def list_active_project_keys(db_path: Path) -> list[str]:
    init_managed_projects_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT project_key
            FROM managed_projects
            WHERE is_active = 1
            ORDER BY project_key ASC
            """
        ).fetchall()
    finally:
        conn.close()
    return [_to_text(row[0]).upper() for row in rows if _to_text(row[0])]


def create_managed_project(db_path: Path, payload: dict) -> dict:
    init_managed_projects_db(db_path)
    project = normalize_managed_project_payload(payload, require_all_fields=True)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO managed_projects (
                project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project["project_key"],
                project["project_name"],
                project["display_name"],
                project["color_hex"],
                project["is_active"],
                now,
                now,
            ),
        )
        conn.commit()
    except sqlite3.IntegrityError:
        raise ValueError(f"Managed project '{project['project_key']}' already exists.")
    finally:
        conn.close()
    out = _load_managed_project_by_key(db_path, project["project_key"])
    if not out:
        raise ValueError("Failed to load created project.")
    return out


def update_managed_project(db_path: Path, project_key: str, payload: dict) -> dict:
    init_managed_projects_db(db_path)
    key = normalize_project_key(project_key)
    existing = _load_managed_project_by_key(db_path, key)
    if not existing:
        raise LookupError(f"Managed project '{key}' not found.")
    normalized = normalize_managed_project_payload(
        {
            **existing,
            **(payload or {}),
            "project_key": key,
        },
        require_all_fields=True,
    )
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE managed_projects
            SET project_name = ?, display_name = ?, color_hex = ?, is_active = ?, updated_at_utc = ?
            WHERE project_key = ?
            """,
            (
                normalized["project_name"],
                normalized["display_name"],
                normalized["color_hex"],
                normalized["is_active"],
                now,
                key,
            ),
        )
        conn.commit()
    finally:
        conn.close()
    out = _load_managed_project_by_key(db_path, key)
    if not out:
        raise LookupError(f"Managed project '{key}' not found.")
    return out


def soft_delete_managed_project(db_path: Path, project_key: str) -> dict:
    init_managed_projects_db(db_path)
    key = normalize_project_key(project_key)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            "UPDATE managed_projects SET is_active = 0, updated_at_utc = ? WHERE project_key = ?",
            (now, key),
        )
        conn.commit()
    finally:
        conn.close()
    if cur.rowcount <= 0:
        raise LookupError(f"Managed project '{key}' not found.")
    out = _load_managed_project_by_key(db_path, key)
    if not out:
        raise LookupError(f"Managed project '{key}' not found.")
    return out


def restore_managed_project(db_path: Path, project_key: str) -> dict:
    init_managed_projects_db(db_path)
    key = normalize_project_key(project_key)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            "UPDATE managed_projects SET is_active = 1, updated_at_utc = ? WHERE project_key = ?",
            (now, key),
        )
        conn.commit()
    finally:
        conn.close()
    if cur.rowcount <= 0:
        raise LookupError(f"Managed project '{key}' not found.")
    out = _load_managed_project_by_key(db_path, key)
    if not out:
        raise LookupError(f"Managed project '{key}' not found.")
    return out


def seed_managed_projects(
    db_path: Path,
    project_keys: list[str],
    project_name_resolver,
) -> dict[str, int]:
    init_managed_projects_db(db_path)
    inserted = 0
    skipped_existing = 0
    for raw_key in project_keys:
        key = normalize_project_key(raw_key)
        existing = _load_managed_project_by_key(db_path, key)
        if existing:
            skipped_existing += 1
            continue
        project_name = _to_text(project_name_resolver(key))
        if not project_name:
            project_name = key
        create_managed_project(
            db_path,
            {
                "project_key": key,
                "project_name": project_name,
                "display_name": project_name,
                "color_hex": deterministic_color_for_project_key(key),
                "is_active": True,
            },
        )
        inserted += 1
    return {"inserted": inserted, "skipped_existing": skipped_existing}


def projects_to_json(projects: list[dict]) -> str:
    return json.dumps(projects, ensure_ascii=True)
