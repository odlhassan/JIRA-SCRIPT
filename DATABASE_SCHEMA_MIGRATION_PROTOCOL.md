# Database Schema Migration Protocol

## Business Logic

Any SQLite schema change in this repo must produce two durable outcomes:

1. A detailed local audit log in `db_schema_changelog.db` describing what changed, why it changed, which files read the changed structure, which tables depend on it, and what the old and new states were.
2. A production-safe migration path that upgrades an older downloaded production database to match the newer local authoritative schema without losing customer data.

The required migration pattern for a modified table is:

1. Rename the production table to `<table>_old`.
2. Create a replacement table with the exact local schema and the original table name.
3. Copy data for compatible columns from `<table>_old` into the replacement table.
4. Validate the migration outcome.
5. Drop `<table>_old` only after the copy succeeds.

Each structural change must be logged with the target DB file, operation type, table name, optional column name, reason, files that read the affected structure, referencing tables, previous state, new state, and the relevant DDL when available.

## Business Cases

This protocol exists because local development usually advances ahead of production. The user may later download an older production database and ask for it to be upgraded to the latest schema while preserving live customer data.

The protocol supports these cases:

- A new feature adds a table or column locally and production must catch up later.
- A column is renamed or its meaning changes, and the migration must preserve old values where possible.
- A table is rebuilt to support new constraints or structure while keeping existing production records.
- A developer needs an audit trail showing when a schema field appeared, why it changed, and what code depends on it.

## Examples

Example 1: add a column

- Local change: add `seal_reason TEXT NOT NULL DEFAULT ''` to `epics_management`.
- Audit log entry: `ADD_COLUMN` for `epics_management.seal_reason`, with the reason, the files that read it, and previous/new state.
- Production migration: rename `epics_management` to `epics_management_old`, create the new `epics_management`, copy common columns, let the new column take its default, then drop `epics_management_old`.

Example 2: rename a column

- Local change: rename `phase_label` to `phase_name` in `epic_phases`.
- Audit log entry: `RENAME_COLUMN` with `old_column_name='phase_label'`, `column_name='phase_name'`, and previous/new type metadata.
- Production migration: rename `epic_phases` to `epic_phases_old`, create the new table, insert with `phase_label AS phase_name`, then drop `epic_phases_old`.

Example 3: remove a column

- Local change: drop a deprecated `legacy_rank` column.
- Audit log entry: `DROP_COLUMN` with the old type and meaning stored in `previous_state_json`.
- Production migration: rebuild the table without `legacy_rank`, copy the surviving columns, then drop the old table.

## Explanations

The local database is treated as the authoritative target schema because it contains the newest feature work. The downloaded production database is treated as an older snapshot that must be transformed forward.

The changelog database is a developer and operator ledger. It answers what changed, why it changed, which code reads it, and what other tables point at it. The migration engine then compares the downloaded production database to the local authoritative database and generates a safe set of SQL steps.

The human dependency is deliberate. A person must download the production database file and place it on the local machine. After that, the agent can inspect both schemas, preview the migration plan, execute the migration, and report the result.

## Front-end UI Fields

This workflow can be operated from CLI today, and the repo also exposes a DB Migration settings route in the server navigation.

| Field / Control | Type | Default | Valid range | What it controls | Example |
|---|---|---|---|---|---|
| Production DB path | file/path input | blank | existing SQLite file | Points to the downloaded production DB to upgrade | `E:\JIRA SCRIPT\downloads\assignee_hours_capacity.prod.db` |
| Local DB path | file/path input | repo DB path | existing SQLite file | Points to the authoritative local schema source | `E:\JIRA SCRIPT\assignee_hours_capacity.db` |
| Changelog DB path | file/path input | `db_schema_changelog.db` | existing or creatable SQLite file | Stores schema audit history and migration runs | `E:\JIRA SCRIPT\db_schema_changelog.db` |
| Plan-only action | button / CLI flag | enabled by choice | boolean | Produces the migration plan without mutating production data | `python db_migration.py --prod ... --local ... --plan-only` |
| Execute migration action | button / CLI action | manual | boolean | Runs rename-create-copy-drop migration steps | `python db_migration.py --prod ... --local ...` |
| Progress log | read-only text / status panel | empty | n/a | Shows per-step diagnostics and errors from `migration_runs` | `Rebuild 'epics_management': +1 added` |

## Script Files

| File | Role |
|---|---|
| `db_schema_changelog.py` | Stores detailed schema change records, schema snapshots, and migration run diagnostics in `db_schema_changelog.db`. |
| `db_migration.py` | Plans and executes production database upgrades by comparing production and local schemas. |
| `report_server.py` | Exposes the DB Migration settings route in the application navigation. |
| `AGENTS.md` | Repo-wide task contract that now requires changelog and migration discipline for future schema changes. |
| `CLAUDE.md` | Persistent repo instructions that define the mandatory schema-change workflow and linked documentation expectations. |
| `.claude/skills/db-schema-migration-discipline/SKILL.md` | Task workflow instructions for future schema-changing work. |

## Dependent & Impacted Files

| File | Relationship |
|---|---|
| `ASSIGNEE_HOURS_CAPACITY.md` | Should be updated whenever the main planning DB schema changes. |
| `docs/report-user-guide/screens/12-epics-planner-tk-estimates.md` | Must reflect Epics Planner table/column changes when schema updates touch that module. |
| Any generator or service file that issues `CREATE TABLE`, `ALTER TABLE`, or reads changed columns | Must be listed in `files_reading_json` and kept aligned with the migration log. |
| Tests covering schema setup or migrations | Must validate the changed structure and any production migration behavior that was modified. |
| `support_center.db` (table `support_issues`) | Standalone, local-only DB for the Support Center report. Recorded in the changelog (ADD_TABLE) and snapshotted, but it is a brand-new DB — no production rename→recreate→copy flow applies. Rebuilt by `support_center_sync.py`; gitignored. The main canonical DBs are untouched/read-only. See `SUPPORT_CENTER_REPORT.md`. |

## Table Schema

The protocol itself uses the following SQLite tables inside `db_schema_changelog.db`.

### `schema_change_log`

| Column | Type | Constraints | Meaning |
|---|---|---|---|
| `id` | INTEGER | primary key autoincrement | Internal row id |
| `change_id` | TEXT | unique not null | Stable UUID for a schema change event |
| `db_file` | TEXT | not null | Target DB file name, such as `assignee_hours_capacity.db` |
| `schema_version` | INTEGER | not null | Monotonic schema version within that DB file |
| `operation` | TEXT | not null | Change type such as `ADD_COLUMN` or `RENAME_TABLE` |
| `table_name` | TEXT | not null default `''` | Table that changed |
| `column_name` | TEXT | not null default `''` | New or current column name |
| `old_column_name` | TEXT | not null default `''` | Previous column name for renames |
| `data_type` | TEXT | not null default `''` | Current data type summary |
| `nullable` | INTEGER | not null default `1` | Whether the new/current column allows null |
| `default_value` | TEXT | not null default `''` | Default value expression |
| `reason` | TEXT | not null default `''` | Why the schema changed |
| `files_reading_json` | TEXT | not null default `'[]'` | JSON array of files that read the changed structure |
| `referencing_tables_json` | TEXT | not null default `'[]'` | JSON array of dependent or foreign-key tables |
| `previous_state_json` | TEXT | not null default `'{}'` | Prior column/table metadata and semantics |
| `new_state_json` | TEXT | not null default `'{}'` | New column/table metadata and semantics |
| `full_ddl` | TEXT | not null default `''` | Relevant DDL statement |
| `changed_at_utc` | TEXT | not null | UTC timestamp for the change |
| `changed_by` | TEXT | not null default `'agent'` | Actor who recorded the change |
| `migration_applied_at_utc` | TEXT | not null default `''` | When production absorbed the change |
| `notes` | TEXT | not null default `''` | Extra operator notes |

### `schema_snapshots`

| Column | Type | Constraints | Meaning |
|---|---|---|---|
| `id` | INTEGER | primary key autoincrement | Internal row id |
| `db_file` | TEXT | not null | Target DB file name |
| `schema_version` | INTEGER | not null, unique with `db_file` | Version represented by the snapshot |
| `snapshot_json` | TEXT | not null | JSON schema dump of tables, columns, foreign keys, and indexes |
| `created_at_utc` | TEXT | not null | UTC snapshot timestamp |

### `migration_runs`

| Column | Type | Constraints | Meaning |
|---|---|---|---|
| `id` | INTEGER | primary key autoincrement | Internal row id |
| `run_id` | TEXT | unique not null | Stable UUID for one migration execution |
| `target_db_file` | TEXT | not null | DB file being upgraded |
| `from_version` | INTEGER | not null default `0` | Starting version known to the changelog |
| `to_version` | INTEGER | not null default `0` | Target version known to the changelog |
| `status` | TEXT | not null default `'running'` | Run state such as `running`, `success`, or `error` |
| `steps_total` | INTEGER | not null default `0` | Planned step count |
| `steps_done` | INTEGER | not null default `0` | Completed step count |
| `log_json` | TEXT | not null default `'[]'` | JSON array of step-by-step diagnostic messages |
| `error_text` | TEXT | not null default `''` | Collapsed error summary |
| `started_at_utc` | TEXT | not null | UTC start time |
| `finished_at_utc` | TEXT | not null default `''` | UTC finish time |

## Data Flow

1. A code change modifies a local SQLite schema in the repo.
2. The developer or agent records the structural change in `db_schema_changelog.py` using `record_change(...)`.
3. The updated local database schema is snapshotted into `schema_snapshots` using `snapshot_current_schema(...)`.
4. The user downloads the older production DB and places it at a local path.
5. `db_migration.py` introspects the production DB and the authoritative local DB.
6. `plan_migration(...)` or `python db_migration.py --plan-only ...` computes table rebuild, create, or drop steps.
7. `migrate_production_db(...)` or `python db_migration.py --prod ... --local ...` executes the plan, logging progress into `migration_runs`.
8. The operator reviews diagnostics and verifies the upgraded production DB now matches the local schema.

## Operator routine

1. Make the local schema change.
2. Record each structural change with `record_change(...)`.
3. Snapshot the updated local schema with `snapshot_current_schema(...)`.
4. Wait for the human-downloaded production DB if it is not yet available.
5. Run a plan preview:

```powershell
python db_migration.py --prod "E:\path\to\production.db" --local "E:\JIRA SCRIPT\assignee_hours_capacity.db" --plan-only
```

6. Execute the migration:

```powershell
python db_migration.py --prod "E:\path\to\production.db" --local "E:\JIRA SCRIPT\assignee_hours_capacity.db"
```

7. Review diagnostics in the CLI output and in `migration_runs` inside `db_schema_changelog.db`.