---
name: db-schema-migration-discipline
description: Mandatory workflow for any SQLite schema change. Use when adding, renaming, deleting, or changing tables, columns, indexes, or constraints so the local changelog DB, schema snapshot, production migration path, and docs stay aligned.
---

# DB Schema Migration Discipline

Use this skill whenever a task changes any SQLite schema in this repo.

## Trigger

- A file adds or changes `CREATE TABLE`, `ALTER TABLE`, index DDL, or schema bootstrap logic.
- A module starts reading a new column or stops reading an old one.
- A local DB now differs structurally from the production DB the user may later download.

## Workflow

### Step 1 — Identify the authoritative local DB and impacted module

- Confirm which DB file is being changed.
- Identify every file that reads or writes the affected table or column.
- Identify related tables, foreign keys, or logical dependents that must be captured in the changelog.

### Step 2 — Make the schema change in local code

- Update the schema-owning Python code, migration helper, or bootstrap logic.
- Keep the local DB as the authoritative target structure.

### Step 3 — Update the local audit trail

- Call `record_change(...)` from `db_schema_changelog.py` for each structural change.
- Capture:
  - `db_file`
  - `operation`
  - `table_name`
  - `column_name` and `old_column_name` when applicable
  - `reason`
  - `files_reading`
  - `referencing_tables`
  - `previous_state`
  - `new_state`
  - `full_ddl`
- Keep `db_schema_changelog.db` up to date as the separate local schema history ledger.

### Step 4 — Snapshot the updated schema

- Run `snapshot_current_schema(...)` for the changed local DB.
- This snapshot becomes the authoritative target for future production upgrade work.

### Step 5 — Prepare the production migration path

- Production migration starts only after the human provides a downloaded production DB file.
- First run a preview:

```powershell
python db_migration.py --prod "E:\path\to\production.db" --local "E:\JIRA SCRIPT\assignee_hours_capacity.db" --plan-only
```

- Then run the real migration:

```powershell
python db_migration.py --prod "E:\path\to\production.db" --local "E:\JIRA SCRIPT\assignee_hours_capacity.db"
```

- The expected table-upgrade pattern is:
  - rename `<table>` to `<table>_old`
  - create the replacement table with the updated local schema
  - copy common-column data into the replacement table
  - drop `<table>_old` only after success

### Step 6 — Keep docs and diagnostics aligned

- Update `DATABASE_SCHEMA_MIGRATION_PROTOCOL.md` when the workflow changes.
- Update the primary module doc and linked docs if the DB change affects business logic or UI.
- If a DB migration UI/route changes, update server wiring, tests, and any related docs.

### Step 7 — Report blockers explicitly

- If the production DB has not yet been downloaded by the user, say so clearly.
- A schema task is not fully complete until the changelog is updated and the production migration path is either executed or blocked only by the missing production DB file.

## Exit Criteria

- Local schema code is updated.
- `db_schema_changelog.db` has change records for the structural updates.
- The updated local schema has been snapshotted.
- The production migration plan is prepared and, when the production DB is available, executed via `db_migration.py`.
- Docs mention the new schema and migration impact.