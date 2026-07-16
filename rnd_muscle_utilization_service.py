from __future__ import annotations

import json
import math
import re
import sqlite3
import uuid
from hashlib import sha1
from datetime import datetime, timezone
from pathlib import Path

from rnd_muscle_utilization_types import (
    DEFAULT_RND_MUSCLE_SKILLS,
    RndMuscleBacklogItem,
    RndMuscleEpic,
    RndMuscleEpicResourceMapping,
    RndMusclePlannerMappingPayload,
    RndMusclePlannerState,
    RndMuscleProjectTab,
    RndMuscleQuickStats,
    RndMuscleResource,
    RndMuscleResourceSkillPayload,
    RndMuscleSkill,
    RndMuscleTeam,
    RndMuscleTeamPayload,
    RndMuscleUtilizationPageState,
)


RND_SOURCE_SCHEMA = "source_db"
DEFAULT_RND_MUSCLE_UTILIZATION_DB = "rnd_muscle_utilization.db"
RND_MUSCLE_TEAM_COLOR_PALETTE: tuple[str, ...] = (
    "#2563eb",
    "#16a34a",
    "#dc2626",
    "#f97316",
    "#ca8a04",
    "#0891b2",
    "#7c3aed",
    "#db2777",
    "#475569",
    "#0f766e",
)
RND_MUSCLE_TABLES: tuple[str, ...] = (
    "rnd_muscle_skills",
    "rnd_muscle_teams",
    "rnd_muscle_team_skills",
    "rnd_muscle_resources",
    "rnd_muscle_resource_skills",
    "rnd_muscle_backlog",
    "rnd_muscle_planner_epics",
    "rnd_muscle_epic_resource_mappings",
)


def _now_utc() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _init_rnd_muscle_utilization_db(settings_db_path: Path, *, seed_defaults: bool = True) -> None:
    settings_db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(settings_db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_skills (
                skill_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                is_default INTEGER NOT NULL DEFAULT 0,
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_teams (
                team_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                color_hex TEXT NOT NULL DEFAULT '#2563eb',
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_team_skills (
                team_id TEXT NOT NULL,
                skill_id TEXT NOT NULL,
                PRIMARY KEY (team_id, skill_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_resources (
                resource_id TEXT PRIMARY KEY,
                display_name TEXT NOT NULL,
                initials TEXT NOT NULL,
                email TEXT NOT NULL DEFAULT '',
                team_id TEXT NOT NULL DEFAULT '',
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_resource_skills (
                resource_id TEXT NOT NULL,
                skill_id TEXT NOT NULL,
                PRIMARY KEY (resource_id, skill_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_backlog (
                epic_key TEXT PRIMARY KEY,
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_planner_epics (
                epic_key TEXT PRIMARY KEY,
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rnd_muscle_epic_resource_mappings (
                epic_key TEXT NOT NULL,
                resource_id TEXT NOT NULL,
                allocation_hours REAL NOT NULL DEFAULT 0.0,
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT '',
                PRIMARY KEY (epic_key, resource_id)
            )
            """
        )
        mapping_cols = {
            str(row[1])
            for row in conn.execute("PRAGMA table_info(rnd_muscle_epic_resource_mappings)").fetchall()
        }
        if "sort_order" not in mapping_cols:
            conn.execute(
                "ALTER TABLE rnd_muscle_epic_resource_mappings ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0"
            )
        planner_count = conn.execute("SELECT COUNT(*) FROM rnd_muscle_planner_epics").fetchone()[0]
        backlog_count = conn.execute("SELECT COUNT(*) FROM rnd_muscle_backlog").fetchone()[0]
        if planner_count == 0 and backlog_count:
            now = _now_utc()
            conn.execute(
                """
                INSERT OR IGNORE INTO rnd_muscle_planner_epics(epic_key, sort_order, created_at_utc, updated_at_utc)
                SELECT epic_key, sort_order, ?, ? FROM rnd_muscle_backlog
                """,
                (now, now),
            )
        # Seed default skills if none exist yet
        existing_skills = conn.execute("SELECT COUNT(*) FROM rnd_muscle_skills").fetchone()[0]
        if seed_defaults and existing_skills == 0:
            now = _now_utc()
            for skill_name in DEFAULT_RND_MUSCLE_SKILLS:
                conn.execute(
                    "INSERT OR IGNORE INTO rnd_muscle_skills(skill_id, name, is_default, created_at_utc, updated_at_utc) VALUES(?,?,1,?,?)",
                    (str(uuid.uuid4()), skill_name, now, now),
                )
        conn.commit()
    finally:
        conn.close()


def migrate_legacy_rnd_muscle_tables(
    legacy_db_path: Path,
    rnd_db_path: Path,
    *,
    drop_legacy: bool = False,
) -> dict[str, object]:
    """Copy legacy rnd_muscle_* tables out of the capacity DB into the dedicated RnD DB."""
    legacy_db_path = Path(legacy_db_path)
    rnd_db_path = Path(rnd_db_path)
    _init_rnd_muscle_utilization_db(rnd_db_path, seed_defaults=False)
    copied: dict[str, int] = {}
    skipped: list[str] = []
    dropped: list[str] = []
    removed_duplicate_skills = 0

    with sqlite3.connect(legacy_db_path) as legacy_conn, sqlite3.connect(rnd_db_path) as rnd_conn:
        legacy_conn.row_factory = sqlite3.Row
        rnd_conn.row_factory = sqlite3.Row
        for table_name in RND_MUSCLE_TABLES:
            if not _table_exists(legacy_conn, table_name):
                skipped.append(table_name)
                continue
            source_cols = _table_columns(legacy_conn, table_name)
            target_cols = _table_column_names(rnd_conn, table_name)
            common_cols = [col for col in target_cols if col in source_cols]
            if not common_cols:
                skipped.append(table_name)
                continue
            col_sql = ", ".join(common_cols)
            placeholder_sql = ", ".join("?" for _ in common_cols)
            rows = legacy_conn.execute(f"SELECT {col_sql} FROM {table_name}").fetchall()
            if rows:
                rnd_conn.executemany(
                    f"INSERT OR REPLACE INTO {table_name} ({col_sql}) VALUES ({placeholder_sql})",
                    [[row[col] for col in common_cols] for row in rows],
                )
            copied[table_name] = len(rows)
        skill_rows = rnd_conn.execute(
            "SELECT rowid, skill_id, lower(name) AS normalized_name FROM rnd_muscle_skills ORDER BY rowid ASC"
        ).fetchall()
        referenced_skill_ids = {
            str(row[0])
            for row in rnd_conn.execute("SELECT skill_id FROM rnd_muscle_team_skills").fetchall()
        }
        referenced_skill_ids.update(
            str(row[0])
            for row in rnd_conn.execute("SELECT skill_id FROM rnd_muscle_resource_skills").fetchall()
        )
        first_rowid_by_name: dict[str, int] = {}
        referenced_rowid_by_name: dict[str, int] = {}
        for row in skill_rows:
            normalized_name = str(row["normalized_name"] or "")
            rowid = int(row["rowid"])
            first_rowid_by_name.setdefault(normalized_name, rowid)
            if str(row["skill_id"]) in referenced_skill_ids:
                referenced_rowid_by_name.setdefault(normalized_name, rowid)
        duplicate_skill_rowids: list[int] = []
        for row in skill_rows:
            normalized_name = str(row["normalized_name"] or "")
            rowid = int(row["rowid"])
            keep_rowid = referenced_rowid_by_name.get(normalized_name, first_rowid_by_name.get(normalized_name, rowid))
            if rowid != keep_rowid and str(row["skill_id"]) not in referenced_skill_ids:
                duplicate_skill_rowids.append(rowid)
        for rowid in duplicate_skill_rowids:
            rnd_conn.execute("DELETE FROM rnd_muscle_skills WHERE rowid = ?", (rowid,))
            removed_duplicate_skills += 1
        rnd_conn.commit()

        if drop_legacy:
            for table_name, source_count in copied.items():
                target_count = rnd_conn.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0]
                if target_count < source_count:
                    raise RuntimeError(
                        f"Refusing to drop {table_name}: target has {target_count} rows after copying {source_count} rows."
                    )
            for table_name in reversed(RND_MUSCLE_TABLES):
                if _table_exists(legacy_conn, table_name):
                    legacy_conn.execute(f"DROP TABLE {table_name}")
                    dropped.append(table_name)
            legacy_conn.commit()

    return {
        "legacy_db_path": str(legacy_db_path),
        "rnd_db_path": str(rnd_db_path),
        "copied": copied,
        "skipped": skipped,
        "dropped": dropped,
        "removed_duplicate_skills": removed_duplicate_skills,
    }


def _priority_int(val: str) -> "int | None":
    text = (val or "").strip().lower()
    mapping = {
        "1": 1,
        "highest": 1,
        "high": 1,
        "2": 2,
        "medium": 2,
        "meidum": 2,
        "3": 3,
        "low": 3,
    }
    return mapping.get(text)


def _is_completed_or_resolved_epic_status(value: object) -> bool:
    """Return whether a canonical Jira epic status is terminal for this planner."""
    normalized = re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()
    return bool({"completed", "resolved"} & set(normalized.split()))


def _latest_canonical_epic_statuses(
    conn: sqlite3.Connection,
    source_schema: str,
    epic_keys: tuple[str, ...],
) -> dict[str, str]:
    """Load statuses only for planner epic keys from the latest canonical snapshot.

    Production canonical_issues can contain millions of historical rows. Filtering
    by run_id first uses its (run_id, issue_key) primary-key index and prevents the
    RnD page request from scanning the full table.
    """
    if not epic_keys or not _table_exists(conn, "canonical_issues", schema=source_schema):
        return {}
    canonical_columns = _table_columns(conn, "canonical_issues", schema=source_schema)
    if not {"issue_key", "status"}.issubset(canonical_columns):
        return {}

    canonical_table = _qualified_table(source_schema, "canonical_issues")
    run_id = ""
    if "run_id" in canonical_columns:
        if _table_exists(conn, "canonical_refresh_state", schema=source_schema):
            state_columns = _table_columns(conn, "canonical_refresh_state", schema=source_schema)
            if "last_success_run_id" in state_columns:
                state_table = _qualified_table(source_schema, "canonical_refresh_state")
                state_row = conn.execute(
                    f"SELECT last_success_run_id FROM {state_table} WHERE id = 1"
                ).fetchone()
                run_id = str(state_row[0] if state_row else "").strip()
        if not run_id:
            run_row = conn.execute(
                f"SELECT run_id FROM {canonical_table} ORDER BY rowid DESC LIMIT 1"
            ).fetchone()
            run_id = str(run_row[0] if run_row else "").strip()
        if not run_id:
            return {}

    statuses: dict[str, str] = {}
    select_issue_type = ", issue_type" if "issue_type" in canonical_columns else ""
    normalized_keys = tuple(dict.fromkeys(key.strip().upper() for key in epic_keys if key.strip()))
    for offset in range(0, len(normalized_keys), 400):
        key_chunk = normalized_keys[offset : offset + 400]
        placeholders = ",".join("?" for _ in key_chunk)
        if "run_id" in canonical_columns:
            sql = (
                f"SELECT issue_key, status{select_issue_type} FROM {canonical_table} "
                f"WHERE run_id = ? AND issue_key IN ({placeholders}) ORDER BY rowid ASC"
            )
            params: tuple[object, ...] = (run_id, *key_chunk)
        else:
            sql = (
                f"SELECT issue_key, status{select_issue_type} FROM {canonical_table} "
                f"WHERE issue_key IN ({placeholders}) ORDER BY rowid ASC"
            )
            params = key_chunk
        for row in conn.execute(sql, params).fetchall():
            if "issue_type" in canonical_columns and str(row["issue_type"] or "").strip().casefold() != "epic":
                continue
            issue_key = str(row["issue_key"] or "").strip().upper()
            if issue_key:
                statuses[issue_key] = str(row["status"] or "")
    return statuses


def _source_schema(rnd_db_path: Path, source_db_path: Path | None) -> str:
    if source_db_path is None:
        return "main"
    try:
        if Path(rnd_db_path).resolve() == Path(source_db_path).resolve():
            return "main"
    except OSError:
        if Path(rnd_db_path) == Path(source_db_path):
            return "main"
    return RND_SOURCE_SCHEMA


def _attach_source_db(conn: sqlite3.Connection, rnd_db_path: Path, source_db_path: Path | None) -> str:
    schema = _source_schema(rnd_db_path, source_db_path)
    if schema == "main":
        return schema
    source = Path(source_db_path or rnd_db_path)
    source.parent.mkdir(parents=True, exist_ok=True)
    conn.execute(f"ATTACH DATABASE ? AS {RND_SOURCE_SCHEMA}", (str(source),))
    return schema


def _connect_rnd_db(rnd_db_path: Path, source_db_path: Path | None = None) -> tuple[sqlite3.Connection, str]:
    conn = sqlite3.connect(rnd_db_path)
    source_schema = _attach_source_db(conn, rnd_db_path, source_db_path)
    return conn, source_schema


def _qualified_table(schema: str, table_name: str) -> str:
    if schema not in {"main", RND_SOURCE_SCHEMA}:
        raise ValueError(f"Unsupported SQLite schema alias: {schema}")
    if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", table_name):
        raise ValueError(f"Unsupported SQLite table name: {table_name}")
    return f"{schema}.{table_name}"


def _table_exists(conn: sqlite3.Connection, table_name: str, *, schema: str = "main") -> bool:
    master_table = _qualified_table(schema, "sqlite_master")
    row = conn.execute(
        f"SELECT 1 FROM {master_table} WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone()
    return row is not None


def _table_columns(conn: sqlite3.Connection, table_name: str, *, schema: str = "main") -> set[str]:
    return {str(col[1]) for col in conn.execute(f"PRAGMA {schema}.table_info({table_name})").fetchall()}


def _table_column_names(conn: sqlite3.Connection, table_name: str, *, schema: str = "main") -> list[str]:
    return [str(col[1]) for col in conn.execute(f"PRAGMA {schema}.table_info({table_name})").fetchall()]


def _resource_id_for_name(display_name: str) -> str:
    normalized = re.sub(r"\s+", " ", (display_name or "").strip())
    digest = sha1(normalized.casefold().encode("utf-8")).hexdigest()[:16]
    return f"canonical-{digest}"


def _initials_for_name(display_name: str) -> str:
    parts = [part for part in re.split(r"\s+", (display_name or "").strip()) if part]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][:1] + parts[-1][:1]).upper()


def _canonical_resource_names(conn: sqlite3.Connection, source_schema: str = "main") -> set[str]:
    names: set[str] = set()
    if _table_exists(conn, "canonical_issues", schema=source_schema):
        issue_cols = _table_columns(conn, "canonical_issues", schema=source_schema)
        if "assignee" in issue_cols:
            issues_table = _qualified_table(source_schema, "canonical_issues")
            for row in conn.execute(f"SELECT DISTINCT assignee FROM {issues_table} WHERE trim(COALESCE(assignee, '')) != ''").fetchall():
                name = str(row[0] or "").strip()
                if name and name.casefold() != "unassigned":
                    names.add(name)
    if _table_exists(conn, "canonical_worklogs", schema=source_schema):
        table_cols = _table_columns(conn, "canonical_worklogs", schema=source_schema)
        worklogs_table = _qualified_table(source_schema, "canonical_worklogs")
        for column_name in ("worklog_author", "issue_assignee"):
            if column_name not in table_cols:
                continue
            for row in conn.execute(f"SELECT DISTINCT {column_name} FROM {worklogs_table} WHERE trim(COALESCE({column_name}, '')) != ''").fetchall():
                name = str(row[0] or "").strip()
                if name and name.casefold() != "unassigned":
                    names.add(name)
    return names


def _sync_canonical_resources(conn: sqlite3.Connection, source_schema: str = "main") -> None:
    now = _now_utc()
    for display_name in sorted(_canonical_resource_names(conn, source_schema), key=str.casefold):
        resource_id = _resource_id_for_name(display_name)
        conn.execute(
            """
            INSERT OR IGNORE INTO rnd_muscle_resources(
                resource_id, display_name, initials, email, team_id, created_at_utc, updated_at_utc
            ) VALUES(?,?,?,?,?,?,?)
            """,
            (resource_id, display_name, _initials_for_name(display_name), "", "", now, now),
        )


def _load_page_state_from_conn(
    conn: sqlite3.Connection,
    *,
    source_schema: str = "main",
    epic_filter_text: str = "",
    epic_project_keys: tuple[str, ...] = (),
) -> RndMuscleUtilizationPageState:
    conn.row_factory = sqlite3.Row
    _sync_canonical_resources(conn, source_schema)
    conn.commit()

    # --- Skills ---
    skill_rows = conn.execute(
        "SELECT skill_id, name, is_default FROM rnd_muscle_skills ORDER BY is_default DESC, lower(name) ASC"
    ).fetchall()
    skills = tuple(
        RndMuscleSkill(skill_id=row["skill_id"], name=row["name"], is_default=bool(row["is_default"]))
        for row in skill_rows
    )

    # --- Teams ---
    team_rows = conn.execute(
        "SELECT team_id, name, color_hex FROM rnd_muscle_teams ORDER BY lower(name) ASC"
    ).fetchall()
    team_skill_rows = conn.execute(
        "SELECT team_id, skill_id FROM rnd_muscle_team_skills"
    ).fetchall()
    team_resource_rows = conn.execute(
        "SELECT team_id, resource_id FROM rnd_muscle_resources WHERE team_id != ''"
    ).fetchall()

    team_skills_by_id: dict[str, list[str]] = {}
    for row in team_skill_rows:
        team_skills_by_id.setdefault(row["team_id"], []).append(row["skill_id"])

    team_resources_by_id: dict[str, list[str]] = {}
    for row in team_resource_rows:
        team_resources_by_id.setdefault(row["team_id"], []).append(row["resource_id"])

    teams = tuple(
        RndMuscleTeam(
            team_id=row["team_id"],
            name=row["name"],
            color_hex=row["color_hex"],
            skill_ids=tuple(team_skills_by_id.get(row["team_id"], [])),
            resource_ids=tuple(team_resources_by_id.get(row["team_id"], [])),
        )
        for row in team_rows
    )

    # --- Resources ---
    resource_rows = conn.execute(
        "SELECT resource_id, display_name, initials, email, team_id FROM rnd_muscle_resources ORDER BY lower(display_name) ASC"
    ).fetchall()
    resource_skill_rows = conn.execute(
        "SELECT resource_id, skill_id FROM rnd_muscle_resource_skills"
    ).fetchall()
    resource_skills_by_id: dict[str, list[str]] = {}
    for row in resource_skill_rows:
        resource_skills_by_id.setdefault(row["resource_id"], []).append(row["skill_id"])
    resignation_by_name: dict[str, str] = {}
    if _table_exists(conn, "performance_resource_resignations", schema=source_schema):
        resignations_table = _qualified_table(source_schema, "performance_resource_resignations")
        for row in conn.execute(
            f"SELECT assignee_name, resignation_date FROM {resignations_table}"
        ).fetchall():
            name = str(row["assignee_name"] or "").strip().casefold()
            if name:
                resignation_by_name[name] = str(row["resignation_date"] or "")

    resources = tuple(
        RndMuscleResource(
            resource_id=row["resource_id"],
            display_name=row["display_name"],
            initials=row["initials"],
            email=row["email"],
            team_id=row["team_id"],
            skill_ids=tuple(resource_skills_by_id.get(row["resource_id"], [])),
            resigned=str(row["display_name"] or "").strip().casefold() in resignation_by_name,
            resignation_date=resignation_by_name.get(str(row["display_name"] or "").strip().casefold(), ""),
        )
        for row in resource_rows
    )

    # --- Epics from Epics Planner rows ---
    epics_table = _qualified_table(source_schema, "epics_management")
    epics_query = f"""
        SELECT em.epic_key, em.epic_name, em.project_key, em.project_name,
               em.priority, em.start_date, em.due_date, em.jira_url
        FROM {epics_table} em
        ORDER BY em.project_key ASC, em.epic_key ASC
    """
    epic_rows = conn.execute(epics_query).fetchall() if _table_exists(conn, "epics_management", schema=source_schema) else []
    latest_epic_status_by_key = _latest_canonical_epic_statuses(
        conn,
        source_schema,
        tuple(str(row["epic_key"] or "").strip().upper() for row in epic_rows),
    )
    completed_or_resolved_epic_keys = {
        epic_key
        for epic_key, status in latest_epic_status_by_key.items()
        if _is_completed_or_resolved_epic_status(status)
    }

    # Budgeted hours from epic_plan column
    plan_values_table = _qualified_table(source_schema, "epics_management_plan_values")
    plan_val_rows = (
        conn.execute(
            f"SELECT epic_key, plan_json FROM {plan_values_table} WHERE column_key = 'epic_plan'"
        ).fetchall()
        if _table_exists(conn, "epics_management_plan_values", schema=source_schema)
        else []
    )
    budgeted_hours_by_key: dict[str, float] = {}
    for row in plan_val_rows:
        try:
            plan = json.loads(row["plan_json"] or "{}")
            man_days = float(plan.get("man_days") or 0.0)
            budgeted_hours_by_key[str(row["epic_key"]).upper()] = man_days * 8.0
        except Exception:
            pass

    # Build all epics
    all_epics_list: list[RndMuscleEpic] = []
    for row in epic_rows:
        epic_key = str(row["epic_key"]).upper()
        if epic_key in completed_or_resolved_epic_keys:
            continue
        all_epics_list.append(
            RndMuscleEpic(
                epic_key=epic_key,
                epic_name=str(row["epic_name"] or ""),
                project_key=str(row["project_key"] or ""),
                project_name=str(row["project_name"] or ""),
                priority=_priority_int(str(row["priority"] or "")),
                budgeted_hours=budgeted_hours_by_key.get(epic_key, 0.0),
                start_date=str(row["start_date"] or ""),
                due_date=str(row["due_date"] or ""),
                jira_url=str(row["jira_url"] or ""),
            )
        )

    # Apply text and project filters for the epic catalog panel
    filtered_epics_list = all_epics_list
    if epic_filter_text:
        lc = epic_filter_text.lower()
        filtered_epics_list = [
            e for e in filtered_epics_list
            if lc in e.epic_key.lower() or lc in e.epic_name.lower() or lc in e.project_key.lower()
        ]
    if epic_project_keys:
        upper_keys = {k.upper() for k in epic_project_keys}
        filtered_epics_list = [e for e in filtered_epics_list if e.project_key.upper() in upper_keys]
    epics = tuple(filtered_epics_list)

    # --- Project tabs based on all epics (not filtered) ---
    project_map: dict[str, str] = {}
    epic_count_by_project: dict[str, int] = {}
    for e in all_epics_list:
        project_map[e.project_key] = e.project_name
        epic_count_by_project[e.project_key] = epic_count_by_project.get(e.project_key, 0) + 1

    total_epic_count = len(all_epics_list)
    project_tabs_list: list[RndMuscleProjectTab] = [
        RndMuscleProjectTab(
            project_key="ALL",
            project_name="All Projects",
            epic_count=total_epic_count,
            is_all_tab=True,
        )
    ]
    for pk in sorted(project_map):
        project_tabs_list.append(
            RndMuscleProjectTab(
                project_key=pk,
                project_name=project_map[pk],
                epic_count=epic_count_by_project[pk],
            )
        )
    project_tabs = tuple(project_tabs_list)

    epics_by_key = {e.epic_key: e for e in all_epics_list}

    def items_from_rows(rows: list[sqlite3.Row] | tuple[sqlite3.Row, ...]) -> tuple[RndMuscleBacklogItem, ...]:
        items: list[RndMuscleBacklogItem] = []
        for row in rows:
            epic_key = str(row["epic_key"]).upper()
            epic = epics_by_key.get(epic_key)
            if epic:
                items.append(
                    RndMuscleBacklogItem(
                        epic_key=epic_key,
                        priority=epic.priority,
                        budgeted_hours=epic.budgeted_hours,
                        start_date=epic.start_date,
                        due_date=epic.due_date,
                        sort_order=int(row["sort_order"]),
                    )
                )
        return tuple(items)

    # --- Planner epics and backlog ---
    planner_epic_rows = conn.execute(
        "SELECT epic_key, sort_order FROM rnd_muscle_planner_epics ORDER BY sort_order ASC, epic_key ASC"
    ).fetchall()
    backlog_rows = conn.execute(
        "SELECT epic_key, sort_order FROM rnd_muscle_backlog ORDER BY sort_order ASC, epic_key ASC"
    ).fetchall()
    planner_epics = items_from_rows(planner_epic_rows)
    backlog = items_from_rows(backlog_rows)

    # --- Mappings ---
    mapping_rows = conn.execute(
        """
        SELECT epic_key, resource_id, allocation_hours, sort_order, created_at_utc, updated_at_utc
        FROM rnd_muscle_epic_resource_mappings
        ORDER BY epic_key ASC, sort_order ASC, resource_id ASC
        """
    ).fetchall()
    mappings = tuple(
        RndMuscleEpicResourceMapping(
            epic_key=str(row["epic_key"]),
            resource_id=str(row["resource_id"]),
            allocation_hours=float(row["allocation_hours"] or 0.0),
            sort_order=int(row["sort_order"] or 0),
            created_at_utc=str(row["created_at_utc"] or ""),
            updated_at_utc=str(row["updated_at_utc"] or ""),
        )
        for row in mapping_rows
    )

    # --- Quick stats ---
    epic_keys_in_mappings = {m.epic_key.upper() for m in mappings}
    resource_ids_in_mappings = {m.resource_id for m in mappings}
    all_resource_ids = {r.resource_id for r in resources}

    resources_associated = len(resource_ids_in_mappings & all_resource_ids)
    resources_not_associated = len(all_resource_ids - resource_ids_in_mappings)
    high_priority_unassigned = sum(
        1 for e in all_epics_list
        if e.priority == 1 and e.epic_key not in epic_keys_in_mappings
    )
    selected_epic_count = len(filtered_epics_list)

    quick_stats = RndMuscleQuickStats(
        resources_associated_with_epics=resources_associated,
        resources_not_yet_associated=resources_not_associated,
        selected_project_epic_count=selected_epic_count,
        high_priority_unassigned_epic_count=high_priority_unassigned,
    )

    planner = RndMusclePlannerState(
        active_project_key="ALL",
        planner_epics=planner_epics,
        backlog=backlog,
        mappings=mappings,
    )

    return RndMuscleUtilizationPageState(
        epics=epics,
        resources=resources,
        teams=teams,
        skills=skills,
        project_tabs=project_tabs,
        quick_stats=quick_stats,
        planner=planner,
    )


def load_rnd_muscle_utilization_page_state(
    settings_db_path: Path,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Load RnD Muscle Utilization state from the feature DB and source planner/canonical DB."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def search_rnd_muscle_epics(
    settings_db_path: Path,
    search_text: str,
    project_keys: tuple[str, ...],
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: filter the left-panel epic catalog by text and selected projects while preserving the same page-state contract used by initial load. Priority should come from Epics Planner Import data when available, falling back to the planner row/default priority rules."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        return _load_page_state_from_conn(
            conn,
            source_schema=source_schema,
            epic_filter_text=(search_text or "").strip(),
            epic_project_keys=project_keys,
        )
    finally:
        conn.close()


def _validate_color_hex(color_hex: str) -> str:
    cleaned = (color_hex or "").strip()
    if not cleaned:
        return "#2563eb"
    if not re.match(r"^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$", cleaned):
        raise ValueError(f"Invalid color_hex '{cleaned}': must be a 3- or 6-digit hex color like #abc or #1a2b3c.")
    normalized = cleaned.lower()
    if normalized not in RND_MUSCLE_TEAM_COLOR_PALETTE:
        raise ValueError("Invalid color_hex: choose one of the 10 supported RnD team colors.")
    return normalized


def save_rnd_muscle_team(
    settings_db_path: Path,
    payload: RndMuscleTeamPayload,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: create or update a manager-defined skillset team, assign its color, associate people, and attach default or custom skills. The final implementation should validate team colors, resource ids, skill ids, duplicate names, and return the refreshed page state."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    team_id = str(payload.get("team_id") or "").strip()
    name = str(payload.get("name") or "").strip()
    if not name:
        raise ValueError("Team name is required.")
    color_hex = _validate_color_hex(str(payload.get("color_hex") or ""))
    raw_skill_ids = payload.get("skill_ids")
    raw_resource_ids = payload.get("resource_ids")
    if raw_skill_ids is not None and not isinstance(raw_skill_ids, (list, tuple)):
        raise ValueError("skill_ids must be an array when provided.")
    if raw_resource_ids is not None and not isinstance(raw_resource_ids, (list, tuple)):
        raise ValueError("resource_ids must be an array when provided.")
    skill_ids: list[str] = [str(s).strip() for s in (raw_skill_ids or []) if str(s).strip()]
    resource_ids: list[str] = [str(r).strip() for r in (raw_resource_ids or []) if str(r).strip()]

    now = _now_utc()
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        _sync_canonical_resources(conn, source_schema)

        # Validate skill_ids
        if skill_ids:
            existing_skill_ids = {
                row["skill_id"] for row in conn.execute("SELECT skill_id FROM rnd_muscle_skills").fetchall()
            }
            unknown = set(skill_ids) - existing_skill_ids
            if unknown:
                raise ValueError(f"Unknown skill_id(s): {sorted(unknown)}")

        # Validate resource_ids
        if resource_ids:
            existing_resource_ids = {
                row["resource_id"] for row in conn.execute("SELECT resource_id FROM rnd_muscle_resources").fetchall()
            }
            unknown_r = set(resource_ids) - existing_resource_ids
            if unknown_r:
                raise ValueError(f"Unknown resource_id(s): {sorted(unknown_r)}")

        if team_id:
            # Update existing team
            existing = conn.execute(
                "SELECT team_id FROM rnd_muscle_teams WHERE team_id = ?", (team_id,)
            ).fetchone()
            if not existing:
                raise ValueError(f"Team not found: {team_id}")
            # Check duplicate name (case-insensitive), excluding self
            dup = conn.execute(
                "SELECT team_id FROM rnd_muscle_teams WHERE lower(name) = lower(?) AND team_id != ?",
                (name, team_id),
            ).fetchone()
            if dup:
                raise ValueError(f"A team named '{name}' already exists.")
            conn.execute(
                "UPDATE rnd_muscle_teams SET name=?, color_hex=?, updated_at_utc=? WHERE team_id=?",
                (name, color_hex, now, team_id),
            )
        else:
            # Create new team
            dup = conn.execute(
                "SELECT team_id FROM rnd_muscle_teams WHERE lower(name) = lower(?)", (name,)
            ).fetchone()
            if dup:
                raise ValueError(f"A team named '{name}' already exists.")
            team_id = str(uuid.uuid4())
            conn.execute(
                "INSERT INTO rnd_muscle_teams(team_id, name, color_hex, created_at_utc, updated_at_utc) VALUES(?,?,?,?,?)",
                (team_id, name, color_hex, now, now),
            )

        if raw_skill_ids is not None:
            conn.execute("DELETE FROM rnd_muscle_team_skills WHERE team_id = ?", (team_id,))
            for sid in skill_ids:
                conn.execute(
                    "INSERT OR IGNORE INTO rnd_muscle_team_skills(team_id, skill_id) VALUES(?,?)",
                    (team_id, sid),
                )

        if raw_resource_ids is not None:
            conn.execute(
                "UPDATE rnd_muscle_resources SET team_id='', updated_at_utc=? WHERE team_id=?",
                (now, team_id),
            )
            for rid in resource_ids:
                conn.execute(
                    "UPDATE rnd_muscle_resources SET team_id=?, updated_at_utc=? WHERE resource_id=?",
                    (team_id, now, rid),
                )

        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def add_rnd_muscle_skill(
    settings_db_path: Path,
    skill_name: str,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: add a manager-defined skill option without duplicating the default skills or existing custom skills, then return the refreshed team/skill configuration state."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    name = (skill_name or "").strip()
    if not name:
        raise ValueError("Skill name is required.")

    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        dup = conn.execute(
            "SELECT skill_id FROM rnd_muscle_skills WHERE lower(name) = lower(?)", (name,)
        ).fetchone()
        if dup:
            raise ValueError(f"A skill named '{name}' already exists.")

        now = _now_utc()
        skill_id = str(uuid.uuid4())
        conn.execute(
            "INSERT INTO rnd_muscle_skills(skill_id, name, is_default, created_at_utc, updated_at_utc) VALUES(?,?,0,?,?)",
            (skill_id, name, now, now),
        )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def save_rnd_muscle_resource_skills(
    settings_db_path: Path,
    payload: RndMuscleResourceSkillPayload,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Persist skills that are explicitly mapped to one resource."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    resource_id = str(payload.get("resource_id") or "").strip()
    if not resource_id:
        raise ValueError("resource_id is required.")
    raw_skill_ids = payload.get("skill_ids")
    if raw_skill_ids is None:
        raw_skill_ids = []
    if not isinstance(raw_skill_ids, (list, tuple)):
        raise ValueError("skill_ids must be an array.")
    skill_ids = [str(skill_id).strip() for skill_id in raw_skill_ids if str(skill_id).strip()]
    if len(skill_ids) != len(set(skill_ids)):
        raise ValueError("skill_ids must not contain duplicates.")

    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        _sync_canonical_resources(conn, source_schema)
        resource = conn.execute(
            "SELECT resource_id FROM rnd_muscle_resources WHERE resource_id = ?",
            (resource_id,),
        ).fetchone()
        if not resource:
            raise ValueError(f"Resource not found: {resource_id}")

        if skill_ids:
            existing_skill_ids = {
                row["skill_id"] for row in conn.execute("SELECT skill_id FROM rnd_muscle_skills").fetchall()
            }
            unknown = set(skill_ids) - existing_skill_ids
            if unknown:
                raise ValueError(f"Unknown skill_id(s): {sorted(unknown)}")

        conn.execute("DELETE FROM rnd_muscle_resource_skills WHERE resource_id = ?", (resource_id,))
        for skill_id in skill_ids:
            conn.execute(
                "INSERT OR IGNORE INTO rnd_muscle_resource_skills(resource_id, skill_id) VALUES(?,?)",
                (resource_id, skill_id),
            )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def add_epic_to_rnd_muscle_backlog(
    settings_db_path: Path,
    epic_key: str,
    sort_order: int | None = None,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: move an epic from the left catalog into the planner backlog so it can be prioritized and mapped to resources. The final implementation should preserve imported priority, budgeted hours, dates, and project context."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    key = (epic_key or "").strip().upper()
    if not key:
        raise ValueError("epic_key is required.")

    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        if not _table_exists(conn, "epics_management", schema=source_schema):
            raise ValueError("Epics Planner table epics_management is not initialized.")
        # Verify epic exists in epics_management
        epic = conn.execute(
            f"SELECT epic_key FROM {_qualified_table(source_schema, 'epics_management')} WHERE upper(epic_key) = ?", (key,)
        ).fetchone()
        if not epic:
            raise ValueError(f"Epic '{key}' not found in Epics Planner.")

        # Determine sort_order: append at end if not specified
        if sort_order is None:
            max_order = conn.execute(
                "SELECT MAX(sort_order) FROM rnd_muscle_backlog"
            ).fetchone()[0]
            sort_order = (int(max_order) + 1) if max_order is not None else 0

        now = _now_utc()
        conn.execute(
            """
            INSERT INTO rnd_muscle_backlog(epic_key, sort_order, created_at_utc, updated_at_utc)
            VALUES(?,?,?,?)
            ON CONFLICT(epic_key) DO UPDATE SET sort_order=excluded.sort_order, updated_at_utc=excluded.updated_at_utc
            """,
            (key, sort_order, now, now),
        )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def remove_epic_from_rnd_muscle_backlog(
    settings_db_path: Path,
    epic_key: str,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    key = (epic_key or "").strip().upper()
    if not key:
        raise ValueError("epic_key is required.")
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.execute("DELETE FROM rnd_muscle_backlog WHERE upper(epic_key) = ?", (key,))
        if cur.rowcount == 0:
            raise ValueError(f"Epic '{key}' is not in the backlog.")
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def add_epic_to_rnd_muscle_planner(
    settings_db_path: Path,
    epic_key: str,
    sort_order: int | None = None,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    key = (epic_key or "").strip().upper()
    if not key:
        raise ValueError("epic_key is required.")

    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        if not _table_exists(conn, "epics_management", schema=source_schema):
            raise ValueError("Epics Planner table epics_management is not initialized.")
        epic = conn.execute(
            f"SELECT epic_key FROM {_qualified_table(source_schema, 'epics_management')} WHERE upper(epic_key) = ?", (key,)
        ).fetchone()
        if not epic:
            raise ValueError(f"Epic '{key}' not found in Epics Planner.")
        if sort_order is None:
            max_order = conn.execute("SELECT MAX(sort_order) FROM rnd_muscle_planner_epics").fetchone()[0]
            sort_order = (int(max_order) + 1) if max_order is not None else 0
        now = _now_utc()
        conn.execute(
            """
            INSERT INTO rnd_muscle_planner_epics(epic_key, sort_order, created_at_utc, updated_at_utc)
            VALUES(?,?,?,?)
            ON CONFLICT(epic_key) DO UPDATE SET sort_order=excluded.sort_order, updated_at_utc=excluded.updated_at_utc
            """,
            (key, sort_order, now, now),
        )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def remove_epic_from_rnd_muscle_planner(
    settings_db_path: Path,
    epic_key: str,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    key = (epic_key or "").strip().upper()
    if not key:
        raise ValueError("epic_key is required.")
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        conn.execute("DELETE FROM rnd_muscle_epic_resource_mappings WHERE upper(epic_key) = ?", (key,))
        cur = conn.execute("DELETE FROM rnd_muscle_planner_epics WHERE upper(epic_key) = ?", (key,))
        if cur.rowcount == 0:
            raise ValueError(f"Epic '{key}' is not in the planner.")
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def save_rnd_muscle_epic_resource_mapping(
    settings_db_path: Path,
    payload: RndMusclePlannerMappingPayload,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: persist resource assignments to an epic on the fly as the manager drops resources onto the hierarchical canvas or adjusts a cluster-view mapping. The final implementation should update quick stats and project-filtered planner state immediately."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    epic_key = str(payload.get("epic_key") or "").strip().upper()
    if not epic_key:
        raise ValueError("epic_key is required.")

    if "resource_ids" not in payload:
        raise ValueError("resource_ids is required.")
    raw_resource_ids = payload.get("resource_ids")
    if not isinstance(raw_resource_ids, (list, tuple)):
        raise ValueError("resource_ids must be an array.")
    resource_ids: list[str] = [str(r).strip() for r in raw_resource_ids if str(r).strip()]
    raw_allocations = payload.get("allocation_hours_by_resource_id") or {}
    if not isinstance(raw_allocations, dict):
        raise ValueError("allocation_hours_by_resource_id must be an object when provided.")
    allocation_hours_by_resource_id: dict[str, float] = {}
    for k, v in raw_allocations.items():
        hours = float(v)
        if not math.isfinite(hours) or hours < 0:
            raise ValueError("allocation hours must be finite and non-negative.")
        allocation_hours_by_resource_id[str(k).strip()] = hours

    now = _now_utc()
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        if not _table_exists(conn, "epics_management", schema=source_schema):
            raise ValueError("Epics Planner table epics_management is not initialized.")
        epic = conn.execute(
            f"SELECT epic_key FROM {_qualified_table(source_schema, 'epics_management')} WHERE upper(epic_key) = ?",
            (epic_key,),
        ).fetchone()
        if not epic:
            raise ValueError(f"Epic '{epic_key}' not found in Epics Planner.")
        if resource_ids:
            existing_resource_ids = {
                row["resource_id"] for row in conn.execute("SELECT resource_id FROM rnd_muscle_resources").fetchall()
            }
            unknown_resource_ids = set(resource_ids) - existing_resource_ids
            if unknown_resource_ids:
                raise ValueError(f"Unknown resource_id(s): {sorted(unknown_resource_ids)}")
        # Remove mappings for resources that are no longer assigned
        existing_mappings = conn.execute(
            "SELECT resource_id FROM rnd_muscle_epic_resource_mappings WHERE epic_key = ?",
            (epic_key,),
        ).fetchall()
        current_resource_ids = {row["resource_id"] for row in existing_mappings}
        to_remove = current_resource_ids - set(resource_ids)
        for rid in to_remove:
            conn.execute(
                "DELETE FROM rnd_muscle_epic_resource_mappings WHERE epic_key=? AND resource_id=?",
                (epic_key, rid),
            )

        # Upsert each resource mapping
        for idx, rid in enumerate(resource_ids):
            hours = allocation_hours_by_resource_id.get(rid, 0.0)
            conn.execute(
                """
                INSERT INTO rnd_muscle_epic_resource_mappings
                    (epic_key, resource_id, allocation_hours, sort_order, created_at_utc, updated_at_utc)
                VALUES(?,?,?,?,?,?)
                ON CONFLICT(epic_key, resource_id) DO UPDATE
                    SET allocation_hours=excluded.allocation_hours,
                        sort_order=excluded.sort_order,
                        updated_at_utc=excluded.updated_at_utc
                """,
                (epic_key, rid, hours, idx, now, now),
            )

        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def reorder_rnd_muscle_backlog(
    settings_db_path: Path,
    epic_keys: list[str] | tuple[str, ...],
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    ordered_keys = [str(key or "").strip().upper() for key in epic_keys if str(key or "").strip()]
    if not ordered_keys:
        raise ValueError("epic_keys must contain at least one epic key.")
    if len(set(ordered_keys)) != len(ordered_keys):
        raise ValueError("epic_keys must not contain duplicates.")

    now = _now_utc()
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        existing_keys = {
            str(row["epic_key"]).upper()
            for row in conn.execute("SELECT epic_key FROM rnd_muscle_backlog").fetchall()
        }
        unknown = set(ordered_keys) - existing_keys
        if unknown:
            raise ValueError(f"Unknown planner epic key(s): {sorted(unknown)}")
        for idx, epic_key in enumerate(ordered_keys):
            conn.execute(
                "UPDATE rnd_muscle_backlog SET sort_order=?, updated_at_utc=? WHERE upper(epic_key)=?",
                (idx, now, epic_key),
            )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def reorder_rnd_muscle_planner_epics(
    settings_db_path: Path,
    epic_keys: list[str] | tuple[str, ...],
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    ordered_keys = [str(key or "").strip().upper() for key in epic_keys if str(key or "").strip()]
    if not ordered_keys:
        raise ValueError("epic_keys must contain at least one epic key.")
    if len(set(ordered_keys)) != len(ordered_keys):
        raise ValueError("epic_keys must not contain duplicates.")

    now = _now_utc()
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        existing_keys = {
            str(row["epic_key"]).upper()
            for row in conn.execute("SELECT epic_key FROM rnd_muscle_planner_epics").fetchall()
        }
        unknown = set(ordered_keys) - existing_keys
        if unknown:
            raise ValueError(f"Unknown planner epic key(s): {sorted(unknown)}")
        for idx, epic_key in enumerate(ordered_keys):
            conn.execute(
                "UPDATE rnd_muscle_planner_epics SET sort_order=?, updated_at_utc=? WHERE upper(epic_key)=?",
                (idx, now, epic_key),
            )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def reorder_rnd_muscle_epic_resources(
    settings_db_path: Path,
    epic_key: str,
    resource_ids: list[str] | tuple[str, ...],
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    _init_rnd_muscle_utilization_db(settings_db_path)
    key = str(epic_key or "").strip().upper()
    if not key:
        raise ValueError("epic_key is required.")
    ordered_resource_ids = [str(rid or "").strip() for rid in resource_ids if str(rid or "").strip()]
    if len(set(ordered_resource_ids)) != len(ordered_resource_ids):
        raise ValueError("resource_ids must not contain duplicates.")

    now = _now_utc()
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        existing = [
            str(row["resource_id"])
            for row in conn.execute(
                "SELECT resource_id FROM rnd_muscle_epic_resource_mappings WHERE upper(epic_key)=?",
                (key,),
            ).fetchall()
        ]
        if set(existing) != set(ordered_resource_ids):
            raise ValueError("resource_ids must match the resources currently mapped to this epic.")
        for idx, resource_id in enumerate(ordered_resource_ids):
            conn.execute(
                """
                UPDATE rnd_muscle_epic_resource_mappings
                SET sort_order=?, updated_at_utc=?
                WHERE upper(epic_key)=? AND resource_id=?
                """,
                (idx, now, key, resource_id),
            )
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()


def list_rnd_muscle_project_tabs(
    settings_db_path: Path,
    selected_project_keys: tuple[str, ...],
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Business logic: build the planner tab strip from configured projects, including the default ALL tab and per-project epic counts for all epics currently participating in the resource planner."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        # Get epics that are actively participating in the resource planner.
        planner_keys = {
            str(row["epic_key"]).upper()
            for row in conn.execute("SELECT epic_key FROM rnd_muscle_planner_epics").fetchall()
        }

        epic_rows = (
            conn.execute(
                f"SELECT epic_key, project_key, project_name FROM {_qualified_table(source_schema, 'epics_management')} ORDER BY project_key ASC"
            ).fetchall()
            if _table_exists(conn, "epics_management", schema=source_schema)
            else []
        )

        project_map: dict[str, str] = {}
        epic_count_by_project: dict[str, int] = {}
        for row in epic_rows:
            ek = str(row["epic_key"]).upper()
            if ek not in planner_keys:
                continue
            pk = str(row["project_key"])
            project_map[pk] = str(row["project_name"])
            epic_count_by_project[pk] = epic_count_by_project.get(pk, 0) + 1

        # Filter to selected projects when provided
        if selected_project_keys:
            upper_selected = {k.upper() for k in selected_project_keys}
            project_map = {k: v for k, v in project_map.items() if k.upper() in upper_selected}
            epic_count_by_project = {k: v for k, v in epic_count_by_project.items() if k.upper() in upper_selected}

        total_in_planner = sum(epic_count_by_project.values())
        project_tabs_list: list[RndMuscleProjectTab] = [
            RndMuscleProjectTab(
                project_key="ALL",
                project_name="All Projects",
                epic_count=total_in_planner,
                is_all_tab=True,
            )
        ]
        for pk in sorted(project_map):
            project_tabs_list.append(
                RndMuscleProjectTab(
                    project_key=pk,
                    project_name=project_map[pk],
                    epic_count=epic_count_by_project[pk],
                )
            )

        page_state = _load_page_state_from_conn(conn, source_schema=source_schema)
        # Return page state with the filtered project_tabs
        return RndMuscleUtilizationPageState(
            report_name=page_state.report_name,
            epics=page_state.epics,
            resources=page_state.resources,
            teams=page_state.teams,
            skills=page_state.skills,
            project_tabs=tuple(project_tabs_list),
            quick_stats=page_state.quick_stats,
            planner=page_state.planner,
        )
    finally:
        conn.close()


def delete_rnd_muscle_team(
    settings_db_path: Path,
    team_id: str,
    source_db_path: Path | None = None,
) -> RndMuscleUtilizationPageState:
    """Delete a manager-defined team by team_id. Clears the team membership from all resources that belonged to this team. Returns the refreshed page state."""
    _init_rnd_muscle_utilization_db(settings_db_path)
    tid = (team_id or "").strip()
    if not tid:
        raise ValueError("team_id is required.")
    conn, source_schema = _connect_rnd_db(settings_db_path, source_db_path)
    try:
        conn.row_factory = sqlite3.Row
        existing = conn.execute(
            "SELECT team_id FROM rnd_muscle_teams WHERE team_id = ?", (tid,)
        ).fetchone()
        if not existing:
            raise ValueError(f"Team not found: {tid}")
        now = _now_utc()
        # Clear team membership for all resources on this team
        conn.execute(
            "UPDATE rnd_muscle_resources SET team_id='', updated_at_utc=? WHERE team_id=?",
            (now, tid),
        )
        # Remove team-skill associations
        conn.execute("DELETE FROM rnd_muscle_team_skills WHERE team_id=?", (tid,))
        # Delete the team
        conn.execute("DELETE FROM rnd_muscle_teams WHERE team_id=?", (tid,))
        conn.commit()
        return _load_page_state_from_conn(conn, source_schema=source_schema)
    finally:
        conn.close()
