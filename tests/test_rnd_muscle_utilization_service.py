from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from rnd_muscle_utilization_service import (
    add_epic_to_rnd_muscle_backlog,
    add_epic_to_rnd_muscle_planner,
    add_rnd_muscle_skill,
    load_rnd_muscle_utilization_page_state,
    migrate_legacy_rnd_muscle_tables,
    remove_epic_from_rnd_muscle_backlog,
    remove_epic_from_rnd_muscle_planner,
    reorder_rnd_muscle_backlog,
    reorder_rnd_muscle_planner_epics,
    reorder_rnd_muscle_epic_resources,
    save_rnd_muscle_epic_resource_mapping,
    save_rnd_muscle_resource_skills,
    save_rnd_muscle_team,
)


def _create_epics_management_table(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE epics_management (
                id TEXT PRIMARY KEY,
                epic_key TEXT NOT NULL,
                project_key TEXT NOT NULL,
                project_name TEXT NOT NULL,
                epic_name TEXT NOT NULL,
                priority TEXT NOT NULL DEFAULT 'Low',
                start_date TEXT NOT NULL DEFAULT '',
                due_date TEXT NOT NULL DEFAULT '',
                jira_url TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE epics_management_plan_values (
                epic_row_id TEXT NOT NULL,
                epic_key TEXT NOT NULL,
                column_key TEXT NOT NULL,
                plan_json TEXT NOT NULL DEFAULT '{}',
                PRIMARY KEY(epic_row_id, column_key)
            )
            """
        )
        conn.executemany(
            """
            INSERT INTO epics_management(
                id, epic_key, project_key, project_name, epic_name, priority, start_date, due_date, jira_url
            ) VALUES(?,?,?,?,?,?,?,?,?)
            """,
            [
                ("row-1", "O2-100", "O2", "OmniConnect", "Highest priority epic", "Highest", "2026-01-01", "2026-01-31", ""),
                ("row-2", "FF-200", "FF", "Fintech Fuel", "Numeric medium epic", "2", "", "", ""),
                ("row-3", "FF-300", "FF", "Fintech Fuel", "Low priority epic", "Low", "", "", ""),
            ],
        )
        conn.execute(
            """
            INSERT INTO epics_management_plan_values(epic_row_id, epic_key, column_key, plan_json)
            VALUES('row-1', 'O2-100', 'epic_plan', '{"man_days": 2}')
            """
        )
        conn.commit()
    finally:
        conn.close()


class RndMuscleUtilizationServiceTests(unittest.TestCase):
    def test_separate_rnd_database_reads_epics_from_source_database(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            source_db_path = root / "assignee_hours_capacity.db"
            rnd_db_path = root / "rnd_muscle_utilization.db"
            _create_epics_management_table(source_db_path)

            state = add_epic_to_rnd_muscle_backlog(rnd_db_path, "O2-100", source_db_path=source_db_path)

            with sqlite3.connect(rnd_db_path) as rnd_conn:
                rnd_tables = {
                    row[0]
                    for row in rnd_conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'rnd_muscle_%'"
                    ).fetchall()
                }
            with sqlite3.connect(source_db_path) as source_conn:
                source_rnd_tables = {
                    row[0]
                    for row in source_conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'rnd_muscle_%'"
                    ).fetchall()
                }

        self.assertEqual([item.epic_key for item in state.planner.backlog], ["O2-100"])
        self.assertIn("rnd_muscle_backlog", rnd_tables)
        self.assertEqual(source_rnd_tables, set())

    def test_migrate_legacy_tables_copies_to_rnd_db_and_can_drop_source_tables(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            legacy_db_path = root / "assignee_hours_capacity.db"
            rnd_db_path = root / "rnd_muscle_utilization.db"
            legacy_state = add_rnd_muscle_skill(legacy_db_path, "Data Platform")
            skill_id = next(skill.skill_id for skill in legacy_state.skills if skill.name == "Data Platform")

            result = migrate_legacy_rnd_muscle_tables(legacy_db_path, rnd_db_path, drop_legacy=True)

            migrated_state = load_rnd_muscle_utilization_page_state(rnd_db_path)
            with sqlite3.connect(legacy_db_path) as legacy_conn:
                legacy_tables = {
                    row[0]
                    for row in legacy_conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'rnd_muscle_%'"
                    ).fetchall()
                }

        self.assertIn("rnd_muscle_skills", result["copied"])
        self.assertIn("rnd_muscle_skills", result["dropped"])
        self.assertTrue(any(skill.skill_id == skill_id for skill in migrated_state.skills))
        self.assertEqual(legacy_tables, set())

    def test_load_on_fresh_db_seeds_default_skills_without_epics_tables(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            state = load_rnd_muscle_utilization_page_state(Path(td) / "settings.db")

        self.assertEqual(state.epics, ())
        self.assertGreaterEqual(len(state.skills), 11)
        self.assertEqual(state.project_tabs[0].project_key, "ALL")
        self.assertEqual(state.project_tabs[0].epic_count, 0)

    def test_priority_normalization_and_filtered_quick_stat_count(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)

            all_state = load_rnd_muscle_utilization_page_state(db_path)
            priority_by_key = {epic.epic_key: epic.priority for epic in all_state.epics}
            self.assertEqual(priority_by_key["O2-100"], 1)
            self.assertEqual(priority_by_key["FF-200"], 2)
            self.assertEqual(priority_by_key["FF-300"], 3)

            from rnd_muscle_utilization_service import search_rnd_muscle_epics

            filtered_state = search_rnd_muscle_epics(db_path, "", ("FF",))
            self.assertEqual(filtered_state.quick_stats.selected_project_epic_count, 2)

    def test_team_update_preserves_omitted_skill_and_resource_assignments(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            state = add_rnd_muscle_skill(db_path, "Data Platform")
            skill_id = next(skill.skill_id for skill in state.skills if skill.name == "Data Platform")
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()

            created = save_rnd_muscle_team(
                db_path,
                {"name": "Backend", "color_hex": "#2563eb", "skill_ids": [skill_id], "resource_ids": ["res-1"]},
            )
            team = next(item for item in created.teams if item.name == "Backend")
            updated = save_rnd_muscle_team(
                db_path,
                {"team_id": team.team_id, "name": "Backend Core", "color_hex": "#16a34a"},
            )
            updated_team = next(item for item in updated.teams if item.team_id == team.team_id)

            self.assertEqual(updated_team.skill_ids, (skill_id,))
            self.assertEqual(updated_team.resource_ids, ("res-1",))

    def test_resource_skill_mapping_persists_direct_skills(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            state = add_rnd_muscle_skill(db_path, "Data Platform")
            skill_id = next(skill.skill_id for skill in state.skills if skill.name == "Data Platform")
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()

            updated = save_rnd_muscle_resource_skills(
                db_path,
                {"resource_id": "res-1", "skill_ids": [skill_id]},
            )
            resource = next(item for item in updated.resources if item.resource_id == "res-1")

        self.assertEqual(resource.skill_ids, (skill_id,))

    def test_resource_skill_mapping_rejects_unknown_resource_and_skill(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            state = load_rnd_muscle_utilization_page_state(db_path)
            skill_id = state.skills[0].skill_id
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()

            with self.assertRaisesRegex(ValueError, "Resource not found"):
                save_rnd_muscle_resource_skills(db_path, {"resource_id": "missing", "skill_ids": [skill_id]})
            with self.assertRaisesRegex(ValueError, "Unknown skill_id"):
                save_rnd_muscle_resource_skills(db_path, {"resource_id": "res-1", "skill_ids": ["missing"]})

    def test_resource_mapping_rejects_unknown_ids_and_negative_hours(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            add_epic_to_rnd_muscle_backlog(db_path, "O2-100")
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()

            with self.assertRaisesRegex(ValueError, "Unknown resource_id"):
                save_rnd_muscle_epic_resource_mapping(
                    db_path,
                    {"epic_key": "O2-100", "resource_ids": ["missing"], "allocation_hours_by_resource_id": {}},
                )
            with self.assertRaisesRegex(ValueError, "finite and non-negative"):
                save_rnd_muscle_epic_resource_mapping(
                    db_path,
                    {"epic_key": "O2-100", "resource_ids": ["res-1"], "allocation_hours_by_resource_id": {"res-1": -1}},
                )
            with self.assertRaisesRegex(ValueError, "not found"):
                save_rnd_muscle_epic_resource_mapping(
                    db_path,
                    {"epic_key": "O2-404", "resource_ids": ["res-1"], "allocation_hours_by_resource_id": {}},
                )

    def test_load_syncs_resources_from_canonical_assignees(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    CREATE TABLE canonical_issues (
                        run_id TEXT NOT NULL,
                        issue_key TEXT NOT NULL,
                        project_key TEXT NOT NULL,
                        issue_type TEXT NOT NULL,
                        summary TEXT NOT NULL,
                        status TEXT NOT NULL,
                        assignee TEXT NOT NULL
                    )
                    """
                )
                conn.execute(
                    """
                    CREATE TABLE canonical_worklogs (
                        run_id TEXT NOT NULL,
                        issue_key TEXT NOT NULL,
                        worklog_author TEXT NOT NULL,
                        issue_assignee TEXT NOT NULL
                    )
                    """
                )
                conn.execute("INSERT INTO canonical_issues VALUES('r1','O2-1','O2','Task','One','Done','Ayesha Khan')")
                conn.execute("INSERT INTO canonical_worklogs VALUES('r1','O2-1','Bilal Ahmed','Ayesha Khan')")
                conn.execute(
                    """
                    CREATE TABLE performance_resource_resignations (
                        assignee_name TEXT PRIMARY KEY,
                        resignation_date TEXT,
                        updated_at TEXT NOT NULL DEFAULT ''
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO performance_resource_resignations(assignee_name, resignation_date, updated_at)
                    VALUES('Ayesha Khan', '2026-06-30', '2026-07-08T00:00:00Z')
                    """
                )
                conn.commit()
            finally:
                conn.close()

            state = load_rnd_muscle_utilization_page_state(db_path)
            names = {resource.display_name for resource in state.resources}
            resource_by_name = {resource.display_name: resource for resource in state.resources}

        self.assertIn("Ayesha Khan", names)
        self.assertIn("Bilal Ahmed", names)
        self.assertTrue(resource_by_name["Ayesha Khan"].resigned)
        self.assertEqual(resource_by_name["Ayesha Khan"].resignation_date, "2026-06-30")
        self.assertFalse(resource_by_name["Bilal Ahmed"].resigned)

    def test_backlog_remove_keeps_planner_mapping_separate(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            add_epic_to_rnd_muscle_backlog(db_path, "O2-100")
            add_epic_to_rnd_muscle_planner(db_path, "O2-100")
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()
            save_rnd_muscle_epic_resource_mapping(
                db_path,
                {"epic_key": "O2-100", "resource_ids": ["res-1"], "allocation_hours_by_resource_id": {"res-1": 4}},
            )

            state = remove_epic_from_rnd_muscle_backlog(db_path, "O2-100")

        self.assertEqual(state.planner.backlog, ())
        self.assertEqual([item.epic_key for item in state.planner.planner_epics], ["O2-100"])
        self.assertEqual([mapping.resource_id for mapping in state.planner.mappings], ["res-1"])

    def test_planner_remove_clears_planner_mapping(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            add_epic_to_rnd_muscle_backlog(db_path, "O2-100")
            add_epic_to_rnd_muscle_planner(db_path, "O2-100")
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES('res-1', 'Hassan Malik', 'HM', 'hassan@example.com', '')
                    """
                )
                conn.commit()
            finally:
                conn.close()
            save_rnd_muscle_epic_resource_mapping(
                db_path,
                {"epic_key": "O2-100", "resource_ids": ["res-1"], "allocation_hours_by_resource_id": {"res-1": 4}},
            )

            state = remove_epic_from_rnd_muscle_planner(db_path, "O2-100")

        self.assertEqual([item.epic_key for item in state.planner.backlog], ["O2-100"])
        self.assertEqual(state.planner.planner_epics, ())
        self.assertEqual(state.planner.mappings, ())

    def test_planner_epic_and_resource_orders_are_persisted(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            _create_epics_management_table(db_path)
            add_epic_to_rnd_muscle_backlog(db_path, "O2-100")
            add_epic_to_rnd_muscle_backlog(db_path, "FF-200")
            add_epic_to_rnd_muscle_planner(db_path, "O2-100")
            add_epic_to_rnd_muscle_planner(db_path, "FF-200")
            conn = sqlite3.connect(db_path)
            try:
                conn.executemany(
                    """
                    INSERT INTO rnd_muscle_resources(resource_id, display_name, initials, email, team_id)
                    VALUES(?,?,?,?,?)
                    """,
                    [
                        ("res-1", "Hassan Malik", "HM", "hassan@example.com", ""),
                        ("res-2", "Ayesha Khan", "AK", "ayesha@example.com", ""),
                    ],
                )
                conn.commit()
            finally:
                conn.close()
            save_rnd_muscle_epic_resource_mapping(
                db_path,
                {"epic_key": "O2-100", "resource_ids": ["res-1", "res-2"], "allocation_hours_by_resource_id": {}},
            )

            backlog_state = reorder_rnd_muscle_backlog(db_path, ["FF-200", "O2-100"])
            epic_state = reorder_rnd_muscle_planner_epics(db_path, ["FF-200", "O2-100"])
            resource_state = reorder_rnd_muscle_epic_resources(db_path, "O2-100", ["res-2", "res-1"])

        self.assertEqual([item.epic_key for item in backlog_state.planner.backlog], ["FF-200", "O2-100"])
        self.assertEqual([item.epic_key for item in epic_state.planner.planner_epics], ["FF-200", "O2-100"])
        self.assertEqual(
            [mapping.resource_id for mapping in resource_state.planner.mappings if mapping.epic_key == "O2-100"],
            ["res-2", "res-1"],
        )
        self.assertEqual(
            [mapping.sort_order for mapping in resource_state.planner.mappings if mapping.epic_key == "O2-100"],
            [0, 1],
        )


if __name__ == "__main__":
    unittest.main()
