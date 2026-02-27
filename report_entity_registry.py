from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


REPORT_ENTITY_GLOBAL_SETTING_KEYS = {
    "planned_leave_min_notice_days",
    "planned_leave_rule_apply_from_date",
    "leave_taken_identification_mode",
    "leave_taken_rule_apply_from_date",
    "rmi_planned_field_resolution",
    "planned_actual_equality_tolerance_hours",
}

DEPRECATED_ENTITY_KEYS = {
    "dispensed_rmi",
    "rmi_dispensing_progress",
}

ALLOWED_FORMULA_FUNCTIONS = {"sum", "count", "min", "max", "average"}


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _normalize_entity_key(value: object) -> str:
    key = _to_text(value).lower()
    if not key:
        raise ValueError("entity_key is required.")
    if not re.match(r"^[a-z0-9_]+$", key):
        raise ValueError(f"Invalid entity_key '{key}'.")
    return key


def _normalize_json_container(value: object, field_name: str) -> list | dict:
    if isinstance(value, (list, dict)):
        return value
    raise ValueError(f"{field_name} must be a JSON object or array.")


def _normalize_formula_expression(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    raise ValueError("formula_expression must be a string.")


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


def _tokenize_formula(expression: str) -> list[tuple[str, str, int]]:
    tokens: list[tuple[str, str, int]] = []
    text = expression or ""
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
            continue
        if ch in "+-*/":
            tokens.append(("OP", ch, i))
            i += 1
            continue
        if ch == "(":
            tokens.append(("LPAREN", ch, i))
            i += 1
            continue
        if ch == ")":
            tokens.append(("RPAREN", ch, i))
            i += 1
            continue
        if ch == ",":
            tokens.append(("COMMA", ch, i))
            i += 1
            continue
        if re.match(r"[A-Za-z_]", ch):
            start = i
            i += 1
            while i < n and re.match(r"[A-Za-z0-9_]", text[i]):
                i += 1
            tokens.append(("IDENT", text[start:i], start))
            continue
        raise ValueError(f"Invalid character '{ch}' at position {i + 1}.")
    tokens.append(("EOF", "", n))
    return tokens


def validate_formula_expression(
    expression: str,
    known_entity_keys: set[str],
    current_entity_key: str = "",
) -> dict:
    text = (expression or "").strip()
    if not text:
        return {"ok": True, "references": []}

    tokens = _tokenize_formula(text)
    idx = 0
    refs: set[str] = set()

    def peek() -> tuple[str, str, int]:
        return tokens[idx]

    def consume(token_type: str | None = None) -> tuple[str, str, int]:
        nonlocal idx
        token = tokens[idx]
        if token_type and token[0] != token_type:
            raise ValueError(f"Expected {token_type.lower()} at position {token[2] + 1}.")
        idx += 1
        return token

    def parse_expr() -> None:
        parse_term()
        while peek()[0] == "OP" and peek()[1] in {"+", "-"}:
            consume("OP")
            parse_term()

    def parse_term() -> None:
        parse_factor()
        while peek()[0] == "OP" and peek()[1] in {"*", "/"}:
            consume("OP")
            parse_factor()

    def parse_factor() -> None:
        token = peek()
        if token[0] == "IDENT":
            ident_token = consume("IDENT")
            ident_lower = ident_token[1].lower()
            if peek()[0] == "LPAREN":
                if ident_lower not in ALLOWED_FORMULA_FUNCTIONS:
                    raise ValueError(
                        f"Unknown function '{ident_token[1]}' at position {ident_token[2] + 1}. "
                        f"Allowed: {', '.join(sorted(ALLOWED_FORMULA_FUNCTIONS))}."
                    )
                consume("LPAREN")
                parse_expr()
                if peek()[0] == "COMMA":
                    raise ValueError(f"Function '{ident_token[1]}' accepts one argument at position {peek()[2] + 1}.")
                consume("RPAREN")
                return
            if ident_lower not in known_entity_keys:
                raise ValueError(f"Unknown entity '{ident_token[1]}' at position {ident_token[2] + 1}.")
            if current_entity_key and ident_lower == current_entity_key.lower():
                raise ValueError(f"Self reference is not allowed for entity '{current_entity_key}'.")
            refs.add(ident_lower)
            return
        if token[0] == "LPAREN":
            consume("LPAREN")
            parse_expr()
            consume("RPAREN")
            return
        raise ValueError(f"Unexpected token at position {token[2] + 1}.")

    parse_expr()
    if peek()[0] != "EOF":
        raise ValueError(f"Unexpected token at position {peek()[2] + 1}.")
    return {"ok": True, "references": sorted(refs)}


def normalize_report_entity_payload(payload: dict) -> dict:
    entity = payload or {}
    out = {
        "entity_key": _normalize_entity_key(entity.get("entity_key")),
        "label": _to_text(entity.get("label")),
        "category": _to_text(entity.get("category")),
        "definition_text": _to_text(entity.get("definition_text")),
        "output_type": _to_text(entity.get("output_type")),
        "identity_level": _to_text(entity.get("identity_level")),
        "source_project_key": _to_text(entity.get("source_project_key")),
        "source_issue_types_json": _normalize_json_container(entity.get("source_issue_types_json"), "source_issue_types_json"),
        "jira_fields_json": _normalize_json_container(entity.get("jira_fields_json"), "jira_fields_json"),
        "selection_rule_json": _normalize_json_container(entity.get("selection_rule_json"), "selection_rule_json"),
        "completeness_rule_json": _normalize_json_container(entity.get("completeness_rule_json"), "completeness_rule_json"),
        "admin_notes": _to_text(entity.get("admin_notes")),
        "is_active": 1 if bool(entity.get("is_active", True)) else 0,
        "formula_expression": _normalize_formula_expression(entity.get("formula_expression")),
        "formula_version": _normalize_formula_version(entity.get("formula_version")),
        "formula_meta_json": _normalize_formula_meta_json(entity.get("formula_meta_json")),
    }
    for req in ("label", "category", "definition_text", "output_type", "identity_level"):
        if not out[req]:
            raise ValueError(f"{req} is required for entity '{out['entity_key']}'.")
    return out


def default_global_settings() -> dict[str, object]:
    return {
        "planned_leave_min_notice_days": 3,
        "planned_leave_rule_apply_from_date": "",
        "leave_taken_identification_mode": "hours",
        "leave_taken_rule_apply_from_date": "",
        "rmi_planned_field_resolution": "name_lookup",
        "planned_actual_equality_tolerance_hours": 0.0,
    }


def normalize_global_settings(payload: dict) -> dict[str, object]:
    source = payload or {}
    unknown = [k for k in source.keys() if k not in REPORT_ENTITY_GLOBAL_SETTING_KEYS]
    if unknown:
        raise ValueError(f"Unknown global settings keys: {', '.join(sorted(unknown))}")
    merged = {**default_global_settings(), **source}
    try:
        merged["planned_leave_min_notice_days"] = max(0, int(merged.get("planned_leave_min_notice_days", 0)))
    except (TypeError, ValueError):
        raise ValueError("planned_leave_min_notice_days must be an integer >= 0.")
    merged["planned_leave_rule_apply_from_date"] = _to_text(merged.get("planned_leave_rule_apply_from_date"))
    merged["leave_taken_rule_apply_from_date"] = _to_text(merged.get("leave_taken_rule_apply_from_date"))
    merged["leave_taken_identification_mode"] = _to_text(merged.get("leave_taken_identification_mode")).lower() or "hours"
    merged["rmi_planned_field_resolution"] = _to_text(merged.get("rmi_planned_field_resolution")).lower() or "name_lookup"
    try:
        merged["planned_actual_equality_tolerance_hours"] = float(merged.get("planned_actual_equality_tolerance_hours", 0.0))
    except (TypeError, ValueError):
        raise ValueError("planned_actual_equality_tolerance_hours must be a number >= 0.")
    if merged["planned_actual_equality_tolerance_hours"] < 0:
        raise ValueError("planned_actual_equality_tolerance_hours must be a number >= 0.")
    merged["planned_actual_equality_tolerance_hours"] = round(merged["planned_actual_equality_tolerance_hours"], 3)
    if merged["leave_taken_identification_mode"] not in {"hours", "status"}:
        raise ValueError("leave_taken_identification_mode must be 'hours' or 'status'.")
    if merged["rmi_planned_field_resolution"] not in {"name_lookup", "field_id", "hybrid"}:
        raise ValueError("rmi_planned_field_resolution must be one of: name_lookup, field_id, hybrid.")
    for key in ("planned_leave_rule_apply_from_date", "leave_taken_rule_apply_from_date"):
        if merged[key] and not re.match(r"^\d{4}-\d{2}-\d{2}$", merged[key]):
            raise ValueError(f"{key} must be empty or ISO date YYYY-MM-DD.")
    return merged


def _entity_seed(
    entity_key: str,
    label: str,
    category: str,
    definition_text: str,
    output_type: str,
    identity_level: str,
    source_project_key: str,
    source_issue_types_json: list,
    jira_fields_json: list,
    selection_rule_json: dict,
    completeness_rule_json: dict,
    admin_notes: str = "",
    formula_expression: str = "",
    formula_version: int = 1,
    formula_meta_json: dict | None = None,
) -> dict:
    return {
        "entity_key": entity_key,
        "label": label,
        "category": category,
        "definition_text": definition_text,
        "output_type": output_type,
        "identity_level": identity_level,
        "source_project_key": source_project_key,
        "source_issue_types_json": source_issue_types_json,
        "jira_fields_json": jira_fields_json,
        "selection_rule_json": selection_rule_json,
        "completeness_rule_json": completeness_rule_json,
        "admin_notes": admin_notes,
        "is_active": 1,
        "formula_expression": formula_expression,
        "formula_version": formula_version,
        "formula_meta_json": formula_meta_json or {},
    }


def report_entity_seed_items() -> list[dict]:
    return [
        _entity_seed("capacity", "Capacity", "core", "Number of business days of a resource in a given time period.", "number_of_business_days", "resource_period", "", [], ["from_date", "to_date", "employee_count", "holiday_dates"], {"method": "capacity_profile_business_days"}, {"required": ["from_date", "to_date"]}, "Definition-only in this phase."),
        _entity_seed("planned_rmi", "Planned RMI", "rmi", 'Epic with "RMI Planned" field set to "Planned"; include full identity bundle.', "epic_identity_bundle", "epic", "", ["Epic"], ["RMI Planned", "summary", "issue_id", "jira_url", "start_date", "due_date", "description", "original_estimate", "product_categorization", "components"], {"rmi_planned_equals": "Planned"}, {"missing_fields_allowed": True}),
        _entity_seed("not_planned_yet_rmi", "Not Planned Yet RMI", "rmi", 'Epic with "RMI Planned" = "Not Planned Yet"; no start/due dates and no logged hours.', "epic_identity_bundle", "epic", "", ["Epic"], ["RMI Planned", "summary", "issue_id", "jira_url", "original_estimate", "start_date", "due_date", "hours_logged"], {"rmi_planned_equals": "Not Planned Yet", "start_date_required": False, "due_date_required": False, "hours_logged_equals": 0}, {"missing_fields_allowed": True}),
        _entity_seed("planned_hours", "Planned Hours", "hours", "Original Estimate in hours.", "hours_duration", "work_item", "", ["Epic", "Story", "Task", "Sub-task"], ["timeoriginalestimate"], {"field": "original_estimate_hours"}, {"missing_fields_allowed": True}),
        _entity_seed("actual_hours", "Actual Hours", "hours", "Hours logged by user on a task.", "hours_duration", "worklog", "", ["Task", "Sub-task", "Bug"], ["timespent", "worklog.timeSpentSeconds"], {"field": "hours_logged"}, {"missing_fields_allowed": True}),
        _entity_seed("log_date", "Log Date", "dates", "Datetime selected for logging Actual Hours.", "datetime", "worklog", "", ["Task", "Sub-task", "Bug"], ["worklog.started"], {"field": "worklog_started"}, {"missing_fields_allowed": True}),
        _entity_seed("planned_dates", "Planned Dates", "dates", "Start date and due date of a work item.", "start_due_date_pair", "work_item", "", ["Epic", "Story", "Task", "Sub-task"], ["start_date", "due_date"], {"fields": ["start_date", "due_date"]}, {"missing_fields_allowed": True}),
        _entity_seed("planned_leaves", "Planned Leaves", "leave", 'In RLT, Leave Type=Planned, status=Resolved, and creation->start >= N days.', "leave_subtask_bundle", "subtask", "RLT", ["Sub-task"], ["Leave Type", "timeoriginalestimate", "created", "start_date", "status", "due_date", "jira_url"], {"leave_type_equals": "Planned", "status_equals": "Resolved", "min_notice_days_setting_key": "planned_leave_min_notice_days", "rule_apply_from_setting_key": "planned_leave_rule_apply_from_date"}, {"missing_fields_allowed": True}),
        _entity_seed("unplanned_leaves", "Unplanned Leaves", "leave", "In RLT, leave classified unplanned when creation->start is less than configured N days.", "leave_subtask_bundle", "subtask", "RLT", ["Sub-task"], ["Leave Type", "timeoriginalestimate", "created", "start_date", "status", "due_date", "jira_url"], {"leave_type_equals": "Planned", "compare": "creation_to_start_less_than_n", "min_notice_days_setting_key": "planned_leave_min_notice_days", "rule_apply_from_setting_key": "planned_leave_rule_apply_from_date"}, {"missing_fields_allowed": True}),
        _entity_seed("planned_leaves_taken", "Planned Leaves Taken", "leave", "Planned Leaves where taken is identified by configured mode (hours or status).", "leave_subtask_bundle", "subtask", "RLT", ["Sub-task"], ["hours_logged", "status"], {"base_entity_key": "planned_leaves", "taken_mode_setting_key": "leave_taken_identification_mode", "rule_apply_from_setting_key": "leave_taken_rule_apply_from_date"}, {"default_taken_mode": "hours"}),
        _entity_seed("unplanned_leaves_taken", "Unplanned Leaves Taken", "leave", "Unplanned Leaves where taken is identified by configured mode (hours or status).", "leave_subtask_bundle", "subtask", "RLT", ["Sub-task"], ["hours_logged", "status"], {"base_entity_key": "unplanned_leaves", "taken_mode_setting_key": "leave_taken_identification_mode", "rule_apply_from_setting_key": "leave_taken_rule_apply_from_date"}, {"default_taken_mode": "hours"}),
        _entity_seed("product_categorization", "Product Categorization", "classification", "Custom Jira field Product Categorization.", "string_or_multi_value", "work_item", "", ["Epic", "Story", "Task", "Sub-task"], ["Product Categorization"], {"field": "product_categorization"}, {"missing_fields_allowed": True}),
        _entity_seed("components", "Components", "classification", "Jira components field with one or more components.", "string_list", "work_item", "", ["Epic", "Story", "Task", "Sub-task"], ["components"], {"field": "components"}, {"missing_fields_allowed": True}),
        _entity_seed("status", "Status", "classification", "Work item status.", "string", "work_item", "", ["Epic", "Story", "Task", "Sub-task", "Bug"], ["status"], {"field": "status"}, {"missing_fields_allowed": True}),
        _entity_seed("rmi", "RMI", "hierarchy", "Epic is called RMI.", "alias", "epic", "", ["Epic"], ["issuetype"], {"issuetype_equals": "Epic"}, {"missing_fields_allowed": True}),
        _entity_seed("phase", "Phase", "hierarchy", "Story under an epic is called Phase of that Epic.", "alias", "story", "", ["Story"], ["parent", "issuetype"], {"issuetype_equals": "Story", "parent_issuetype_equals": "Epic"}, {"missing_fields_allowed": True}),
        _entity_seed("activity", "Activity", "hierarchy", "Subtask or bug subtask under a story.", "alias", "subtask", "", ["Sub-task", "Bug Subtask"], ["parent", "issuetype"], {"issuetype_in": ["Sub-task", "Bug Subtask"], "parent_issuetype_equals": "Story"}, {"missing_fields_allowed": True}),
    ]


def _ensure_table_column(conn: sqlite3.Connection, table_name: str, column_name: str, definition_sql: str) -> None:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    existing = {str(row[1]) for row in rows}
    if column_name in existing:
        return
    conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition_sql}")


def init_report_entities_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS report_entity_definitions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_key TEXT NOT NULL UNIQUE,
                label TEXT NOT NULL,
                category TEXT NOT NULL,
                definition_text TEXT NOT NULL,
                output_type TEXT NOT NULL,
                identity_level TEXT NOT NULL,
                source_project_key TEXT,
                source_issue_types_json TEXT NOT NULL,
                jira_fields_json TEXT NOT NULL,
                selection_rule_json TEXT NOT NULL,
                completeness_rule_json TEXT NOT NULL,
                admin_notes TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                formula_expression TEXT NOT NULL DEFAULT '',
                formula_version INTEGER NOT NULL DEFAULT 1,
                formula_meta_json TEXT NOT NULL DEFAULT '{}',
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS report_entity_global_settings (
                key TEXT PRIMARY KEY,
                value_json TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        _ensure_table_column(conn, "report_entity_definitions", "formula_expression", "TEXT NOT NULL DEFAULT ''")
        _ensure_table_column(conn, "report_entity_definitions", "formula_version", "INTEGER NOT NULL DEFAULT 1")
        _ensure_table_column(conn, "report_entity_definitions", "formula_meta_json", "TEXT NOT NULL DEFAULT '{}'")
        conn.commit()
    finally:
        conn.close()


def seed_report_entities_if_empty(db_path: Path) -> None:
    init_report_entities_db(db_path)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        if DEPRECATED_ENTITY_KEYS:
            placeholders = ", ".join("?" for _ in DEPRECATED_ENTITY_KEYS)
            conn.execute(
                f"DELETE FROM report_entity_definitions WHERE entity_key IN ({placeholders})",
                tuple(sorted(DEPRECATED_ENTITY_KEYS)),
            )
        count = int(conn.execute("SELECT COUNT(*) AS c FROM report_entity_definitions").fetchone()["c"] or 0)
        if count == 0:
            for seed in report_entity_seed_items():
                entity = normalize_report_entity_payload(seed)
                conn.execute(
                    """
                    INSERT INTO report_entity_definitions (
                        entity_key, label, category, definition_text, output_type, identity_level, source_project_key,
                        source_issue_types_json, jira_fields_json, selection_rule_json, completeness_rule_json,
                        admin_notes, is_active, formula_expression, formula_version, formula_meta_json, created_at_utc, updated_at_utc
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        entity["entity_key"],
                        entity["label"],
                        entity["category"],
                        entity["definition_text"],
                        entity["output_type"],
                        entity["identity_level"],
                        entity["source_project_key"],
                        json.dumps(entity["source_issue_types_json"]),
                        json.dumps(entity["jira_fields_json"]),
                        json.dumps(entity["selection_rule_json"]),
                        json.dumps(entity["completeness_rule_json"]),
                        entity["admin_notes"],
                        entity["is_active"],
                        entity["formula_expression"],
                        entity["formula_version"],
                        json.dumps(entity["formula_meta_json"]),
                        now,
                        now,
                    ),
                )
        defaults = normalize_global_settings(default_global_settings())
        for key, value in defaults.items():
            row = conn.execute("SELECT key FROM report_entity_global_settings WHERE key = ?", (key,)).fetchone()
            if row:
                continue
            conn.execute(
                "INSERT INTO report_entity_global_settings (key, value_json, updated_at_utc) VALUES (?, ?, ?)",
                (key, json.dumps(value), now),
            )
        conn.commit()
    finally:
        conn.close()


def load_report_entities(db_path: Path) -> list[dict]:
    seed_report_entities_if_empty(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT id, entity_key, label, category, definition_text, output_type, identity_level, source_project_key,
                   source_issue_types_json, jira_fields_json, selection_rule_json, completeness_rule_json,
                   admin_notes, is_active, formula_expression, formula_version, formula_meta_json, created_at_utc, updated_at_utc
            FROM report_entity_definitions
            ORDER BY lower(label) ASC, entity_key ASC
            """
        ).fetchall()
    finally:
        conn.close()
    out: list[dict] = []
    for row in rows:
        try:
            formula_meta_json = json.loads(_to_text(row["formula_meta_json"]) or "{}")
            if not isinstance(formula_meta_json, dict):
                formula_meta_json = {}
        except json.JSONDecodeError:
            formula_meta_json = {}
        out.append(
            {
                "id": int(row["id"]),
                "entity_key": _to_text(row["entity_key"]),
                "label": _to_text(row["label"]),
                "category": _to_text(row["category"]),
                "definition_text": _to_text(row["definition_text"]),
                "output_type": _to_text(row["output_type"]),
                "identity_level": _to_text(row["identity_level"]),
                "source_project_key": _to_text(row["source_project_key"]),
                "source_issue_types_json": json.loads(_to_text(row["source_issue_types_json"]) or "[]"),
                "jira_fields_json": json.loads(_to_text(row["jira_fields_json"]) or "[]"),
                "selection_rule_json": json.loads(_to_text(row["selection_rule_json"]) or "{}"),
                "completeness_rule_json": json.loads(_to_text(row["completeness_rule_json"]) or "{}"),
                "admin_notes": _to_text(row["admin_notes"]),
                "is_active": bool(int(row["is_active"] or 0)),
                "formula_expression": _to_text(row["formula_expression"]),
                "formula_version": int(row["formula_version"] or 1),
                "formula_meta_json": formula_meta_json,
                "created_at_utc": _to_text(row["created_at_utc"]),
                "updated_at_utc": _to_text(row["updated_at_utc"]),
            }
        )
    return out


def load_report_entity_global_settings(db_path: Path) -> dict[str, object]:
    seed_report_entities_if_empty(db_path)
    out = normalize_global_settings(default_global_settings())
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute("SELECT key, value_json FROM report_entity_global_settings").fetchall()
    finally:
        conn.close()
    for row in rows:
        key = _to_text(row["key"])
        if key not in REPORT_ENTITY_GLOBAL_SETTING_KEYS:
            continue
        try:
            out[key] = json.loads(_to_text(row["value_json"]) or "null")
        except json.JSONDecodeError:
            continue
    return normalize_global_settings(out)


def save_report_entities(db_path: Path, entities: list[dict]) -> list[dict]:
    if not isinstance(entities, list) or not entities:
        raise ValueError("entities must be a non-empty array.")
    normalized = [normalize_report_entity_payload(item) for item in entities]
    keys: set[str] = set()
    for item in normalized:
        if item["entity_key"] in keys:
            raise ValueError(f"Duplicate entity_key: {item['entity_key']}")
        keys.add(item["entity_key"])
    known_keys = {item["entity_key"] for item in normalized}
    for item in normalized:
        try:
            formula_validation = validate_formula_expression(
                item["formula_expression"],
                known_entity_keys=known_keys,
                current_entity_key=item["entity_key"],
            )
        except ValueError as err:
            raise ValueError(f"Invalid formula_expression for entity '{item['entity_key']}': {err}")
        item["formula_meta_json"]["references"] = formula_validation.get("references", [])
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        for entity in normalized:
            conn.execute(
                """
                INSERT INTO report_entity_definitions (
                    entity_key, label, category, definition_text, output_type, identity_level, source_project_key,
                    source_issue_types_json, jira_fields_json, selection_rule_json, completeness_rule_json,
                    admin_notes, is_active, formula_expression, formula_version, formula_meta_json, created_at_utc, updated_at_utc
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(entity_key) DO UPDATE SET
                    label=excluded.label,
                    category=excluded.category,
                    definition_text=excluded.definition_text,
                    output_type=excluded.output_type,
                    identity_level=excluded.identity_level,
                    source_project_key=excluded.source_project_key,
                    source_issue_types_json=excluded.source_issue_types_json,
                    jira_fields_json=excluded.jira_fields_json,
                    selection_rule_json=excluded.selection_rule_json,
                    completeness_rule_json=excluded.completeness_rule_json,
                    admin_notes=excluded.admin_notes,
                    is_active=excluded.is_active,
                    formula_expression=excluded.formula_expression,
                    formula_version=excluded.formula_version,
                    formula_meta_json=excluded.formula_meta_json,
                    updated_at_utc=excluded.updated_at_utc
                """,
                (
                    entity["entity_key"],
                    entity["label"],
                    entity["category"],
                    entity["definition_text"],
                    entity["output_type"],
                    entity["identity_level"],
                    entity["source_project_key"],
                    json.dumps(entity["source_issue_types_json"]),
                    json.dumps(entity["jira_fields_json"]),
                    json.dumps(entity["selection_rule_json"]),
                    json.dumps(entity["completeness_rule_json"]),
                    entity["admin_notes"],
                    entity["is_active"],
                    entity["formula_expression"],
                    entity["formula_version"],
                    json.dumps(entity["formula_meta_json"]),
                    now,
                    now,
                ),
            )
        conn.commit()
    finally:
        conn.close()
    return load_report_entities(db_path)


def save_report_entity_global_settings(db_path: Path, settings: dict) -> dict[str, object]:
    normalized = normalize_global_settings(settings)
    now = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        for key, value in normalized.items():
            conn.execute(
                """
                INSERT INTO report_entity_global_settings (key, value_json, updated_at_utc)
                VALUES (?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    value_json=excluded.value_json,
                    updated_at_utc=excluded.updated_at_utc
                """,
                (key, json.dumps(value), now),
            )
        conn.commit()
    finally:
        conn.close()
    return load_report_entity_global_settings(db_path)


def reset_report_entities_to_defaults(db_path: Path) -> dict[str, object]:
    init_report_entities_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("DELETE FROM report_entity_definitions")
        conn.execute("DELETE FROM report_entity_global_settings")
        conn.commit()
    finally:
        conn.close()
    seed_report_entities_if_empty(db_path)
    return {
        "entities": load_report_entities(db_path),
        "global_settings": load_report_entity_global_settings(db_path),
        "source": "db",
    }
