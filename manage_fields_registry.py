from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from report_entity_registry import load_report_entities, validate_formula_expression


ALLOWED_MANAGED_FIELD_DATA_TYPES = {"number", "text", "date", "boolean"}


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _normalize_field_key(value: object) -> str:
    key = _to_text(value).lower()
    if not key:
        raise ValueError("field_key is required.")
    if not re.match(r"^[a-z0-9_]+$", key):
        raise ValueError(f"Invalid field_key '{key}'.")
    return key


def _slugify_field_key_from_label(label: str) -> str:
    text = _to_text(label).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "field"


def _next_available_field_key(conn: sqlite3.Connection, base_key: str) -> str:
    key = _normalize_field_key(base_key)
    suffix = 1
    candidate = key
    while True:
        row = conn.execute("SELECT 1 FROM managed_fields WHERE field_key = ?", (candidate,)).fetchone()
        if not row:
            return candidate
        suffix += 1
        candidate = f"{key}_{suffix}"


def _normalize_formula_version(value: object) -> int:
    if value is None or value == "":
        return 1
    try:
        out = int(value)
    except (TypeError, ValueError):
        raise ValueError("formula_version must be an integer >= 1.")
    if out < 1:
        raise ValueError("formula_version must be an integer >= 1.")
    return out


def _normalize_formula_meta_json(value: object) -> dict:
    if value is None or value == "":
        return {}
    if isinstance(value, dict):
        return value
    raise ValueError("formula_meta_json must be a JSON object.")


def normalize_managed_field_payload(payload: dict, require_key: bool = True) -> dict:
    raw = payload or {}
    field_key = _normalize_field_key(raw.get("field_key")) if require_key else _to_text(raw.get("field_key")).lower()
    label = _to_text(raw.get("label"))
    description = _to_text(raw.get("description"))
    data_type = _to_text(raw.get("data_type")).lower()
    formula_expression = _to_text(raw.get("formula_expression"))
    formula_version = _normalize_formula_version(raw.get("formula_version"))
    formula_meta_json = _normalize_formula_meta_json(raw.get("formula_meta_json"))
    is_active = 1 if bool(raw.get("is_active", True)) else 0

    if require_key and not field_key:
        raise ValueError("field_key is required.")
    if not label:
        raise ValueError("label is required.")
    if not data_type:
        raise ValueError("data_type is required.")
    if data_type not in ALLOWED_MANAGED_FIELD_DATA_TYPES:
        raise ValueError(
            "data_type must be one of: " + ", ".join(sorted(ALLOWED_MANAGED_FIELD_DATA_TYPES)) + "."
        )
    return {
        "field_key": field_key,
        "label": label,
        "description": description,
        "data_type": data_type,
        "formula_expression": formula_expression,
        "formula_version": formula_version,
        "formula_meta_json": formula_meta_json,
        "is_active": is_active,
    }


def init_manage_fields_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS managed_fields (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                field_key TEXT NOT NULL UNIQUE,
                label TEXT NOT NULL,
                description TEXT NOT NULL,
                data_type TEXT NOT NULL,
                formula_expression TEXT NOT NULL DEFAULT '',
                formula_version INTEGER NOT NULL DEFAULT 1,
                formula_meta_json TEXT NOT NULL DEFAULT '{}',
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _row_to_field(row: sqlite3.Row) -> dict:
    try:
        meta = json.loads(_to_text(row["formula_meta_json"]) or "{}")
        if not isinstance(meta, dict):
            meta = {}
    except json.JSONDecodeError:
        meta = {}
    return {
        "id": int(row["id"]),
        "field_key": _to_text(row["field_key"]),
        "label": _to_text(row["label"]),
        "description": _to_text(row["description"]),
        "data_type": _to_text(row["data_type"]),
        "formula_expression": _to_text(row["formula_expression"]),
        "formula_version": int(row["formula_version"] or 1),
        "formula_meta_json": meta,
        "is_active": bool(int(row["is_active"] or 0)),
        "created_at_utc": _to_text(row["created_at_utc"]),
        "updated_at_utc": _to_text(row["updated_at_utc"]),
    }


def _load_managed_field_by_key(db_path: Path, field_key: str) -> dict | None:
    init_manage_fields_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT id, field_key, label, description, data_type,
                   formula_expression, formula_version, formula_meta_json,
                   is_active, created_at_utc, updated_at_utc
            FROM managed_fields
            WHERE field_key = ?
            """,
            (field_key,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    return _row_to_field(row)


def load_manage_fields(db_path: Path, include_inactive: bool = False) -> list[dict]:
    init_manage_fields_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        if include_inactive:
            rows = conn.execute(
                """
                SELECT id, field_key, label, description, data_type,
                       formula_expression, formula_version, formula_meta_json,
                       is_active, created_at_utc, updated_at_utc
                FROM managed_fields
                ORDER BY lower(label) ASC, field_key ASC
                """
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT id, field_key, label, description, data_type,
                       formula_expression, formula_version, formula_meta_json,
                       is_active, created_at_utc, updated_at_utc
                FROM managed_fields
                WHERE is_active = 1
                ORDER BY lower(label) ASC, field_key ASC
                """
            ).fetchall()
    finally:
        conn.close()
    return [_row_to_field(row) for row in rows]


def _validate_field_formula_against_entities(db_path: Path, field: dict) -> dict:
    known_entity_keys = {item["entity_key"] for item in load_report_entities(db_path)}
    try:
        formula_validation = validate_formula_expression(
            field["formula_expression"],
            known_entity_keys=known_entity_keys,
            current_entity_key="",
        )
    except ValueError as err:
        raise ValueError(f"Invalid formula_expression for field '{field['field_key']}': {err}")
    meta = dict(field["formula_meta_json"])
    meta["references"] = formula_validation.get("references", [])
    field["formula_meta_json"] = meta
    return field


def create_manage_field(db_path: Path, payload: dict) -> dict:
    init_manage_fields_db(db_path)
    raw = payload or {}
    field = normalize_managed_field_payload({**raw, "field_key": "__system_managed__"}, require_key=True)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        # Field key is always system-managed and generated from label on create.
        base_key = _slugify_field_key_from_label(field["label"])
        field["field_key"] = _next_available_field_key(conn, base_key)
        field = _validate_field_formula_against_entities(db_path, field)
        conn.execute(
            """
            INSERT INTO managed_fields (
                field_key, label, description, data_type,
                formula_expression, formula_version, formula_meta_json,
                is_active, created_at_utc, updated_at_utc
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                field["field_key"],
                field["label"],
                field["description"],
                field["data_type"],
                field["formula_expression"],
                field["formula_version"],
                json.dumps(field["formula_meta_json"]),
                field["is_active"],
                now,
                now,
            ),
        )
        conn.commit()
    finally:
        conn.close()
    out = _load_managed_field_by_key(db_path, field["field_key"])
    if not out:
        raise ValueError("Failed to load created field.")
    return out


def update_manage_field(db_path: Path, field_key: str, payload: dict) -> dict:
    init_manage_fields_db(db_path)
    normalized_key = _normalize_field_key(field_key)
    existing = _load_managed_field_by_key(db_path, normalized_key)
    if not existing:
        raise LookupError(f"Managed field '{normalized_key}' not found.")

    incoming = normalize_managed_field_payload({**existing, **(payload or {}), "field_key": normalized_key}, require_key=True)
    incoming = _validate_field_formula_against_entities(db_path, incoming)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            UPDATE managed_fields
            SET label = ?, description = ?, data_type = ?,
                formula_expression = ?, formula_version = ?, formula_meta_json = ?,
                is_active = ?, updated_at_utc = ?
            WHERE field_key = ?
            """,
            (
                incoming["label"],
                incoming["description"],
                incoming["data_type"],
                incoming["formula_expression"],
                incoming["formula_version"],
                json.dumps(incoming["formula_meta_json"]),
                incoming["is_active"],
                now,
                normalized_key,
            ),
        )
        conn.commit()
    finally:
        conn.close()
    out = _load_managed_field_by_key(db_path, normalized_key)
    if not out:
        raise LookupError(f"Managed field '{normalized_key}' not found.")
    return out


def soft_delete_manage_field(db_path: Path, field_key: str) -> dict:
    init_manage_fields_db(db_path)
    normalized_key = _normalize_field_key(field_key)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            "UPDATE managed_fields SET is_active = 0, updated_at_utc = ? WHERE field_key = ?",
            (now, normalized_key),
        )
        conn.commit()
    finally:
        conn.close()
    if cur.rowcount <= 0:
        raise LookupError(f"Managed field '{normalized_key}' not found.")
    out = _load_managed_field_by_key(db_path, normalized_key)
    if not out:
        raise LookupError(f"Managed field '{normalized_key}' not found.")
    return out


def restore_manage_field(db_path: Path, field_key: str) -> dict:
    init_manage_fields_db(db_path)
    normalized_key = _normalize_field_key(field_key)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(
            "UPDATE managed_fields SET is_active = 1, updated_at_utc = ? WHERE field_key = ?",
            (now, normalized_key),
        )
        conn.commit()
    finally:
        conn.close()
    if cur.rowcount <= 0:
        raise LookupError(f"Managed field '{normalized_key}' not found.")
    out = _load_managed_field_by_key(db_path, normalized_key)
    if not out:
        raise LookupError(f"Managed field '{normalized_key}' not found.")
    return out
