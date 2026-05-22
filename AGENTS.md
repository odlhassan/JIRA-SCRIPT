# AGENTS.md

## Final response requirement
- End every completed task with two short sections in this order:
  - `How to experience latest changes on live localhost`
  - `How to test locally`
- Before writing those sections, inspect project-native sources such as `AGENTS.md`, `README.md`, `package.json`, test files, helper scripts, and existing CLI/docs in the repo instead of guessing commands, ports, or URLs.
- When verification is feasible in the current environment, execute the narrowest relevant local CLI commands yourself before completing the task. Prefer focused checks over broad expensive runs unless the change requires broader coverage.
- If the change touches HTML, UI, templates, generated reports, or served assets, update coupled JavaScript, CSS, generator, server, test, and relevant `.md` documentation files instead of changing HTML alone.
- In the final response, include the exact CLI commands you executed and separately list any additional local commands the user can re-run on the same machine.
- If a local server run is applicable, include the exact command to start the server, the exact localhost URL or route to inspect, ordered manual steps, and the expected visible behavior.
- If a command could not be executed, say so briefly and explain why.
- If no localhost run is applicable, explicitly state `How to experience latest changes on live localhost: Not applicable for this change`.
- If no local server run is applicable, explicitly state `How to test locally: Not applicable for this change`.

## Git push to GitHub

- When the user asks to **push to GitHub** (or equivalent) **without naming a branch**, use **`main`**: `git checkout main`, commit there if needed, then `git push origin main` so Azure deploy can run. Do not default to `backup/*` or other branches unless the user explicitly asks for that branch.

## Database schema change protocol

- Any SQLite schema change is incomplete until the local audit trail and production migration path are updated.
- Use `db_schema_changelog.py` to keep a separate local changelog database (`db_schema_changelog.db`) with, at minimum, the changed table/column, operation, reason, files that read the column, referencing tables, and previous/new state.
- After changing a local database structure, snapshot the authoritative local schema so the repo has a current target definition for later production migration work.
- Production migration depends on a human-downloaded production DB file. Once the user provides that DB locally, use `db_migration.py` to plan and execute a rename-create-copy-drop migration that preserves customer data:
  - rename modified production tables to `_old`
  - create replacement tables with the updated local structure
  - migrate shared-column data into the replacement tables
  - drop the `_old` tables after successful copy
- If a schema task finishes before the production DB is available, explicitly report that the changelog was updated but the production migration run is pending the user-provided DB file.
- If the migration workflow, diagnostics, or UI changes, update the related Python, UI, tests, and `.md` files together.
