# Azure App Service Deployment

This project can run on Azure App Service for Linux as a Python 3.11 web app.

## GitHub Actions → production

- **Live URL:** `https://epreporting.azurewebsites.net/`
- Workflow: `.github/workflows/azure-appservice-deploy.yml` (runs on push to `main` or `master`, and manual `workflow_dispatch`).
- The deploy ZIP is built from the **GitHub checkout** after staging (excludes tests, `node_modules`, `handover`, `offline_bundles`, root `backup/`, etc.). **Commit and push every file the running app depends on** — anything not in Git will not ship. Do not rely on unpushed local builds.
- The workflow runs `pip install -r requirements.txt --target staging/.python_packages/lib/site-packages` before zipping. This makes production dependency imports independent of whether Azure creates `/home/site/wwwroot/antenv` during deployment. It also marks `startup.txt` executable before creating the ZIP so App Service can launch it from a read-only `WEBSITE_RUN_FROM_PACKAGE` mount.
- Before uploading the ZIP, the workflow verifies that required startup files and helpers are present in `staging/`, including `startup.txt`, `wsgi.py`, `report_server.py`, `db_journal_mode.py`, `requirements.txt`, and the vendored Gunicorn package.
- **Backups:** do not commit local backup trees or `*.backup` files (see `.gitignore`). Do not expect backup-only branches to drive production unless you change the workflow deliberately.

## Why a startup command is required

Azure App Service auto-detects Flask only when the app entrypoint is `app.py` or `application.py` with an `app` callable.

This repo exposes the Flask app from `wsgi.py`, so configure Azure to use the included executable `startup.txt` script.
The startup script runs Gunicorn on Azure's injected `$PORT` with a single worker:

```text
startup.txt
```

Inside `startup.txt`, the script prepends `.python_packages/lib/site-packages` to
`PYTHONPATH` and then execs Gunicorn:

```sh
export PYTHONPATH="$(pwd)/.python_packages/lib/site-packages:${PYTHONPATH:-}"
exec python -m gunicorn --bind=0.0.0.0:${PORT:-8000} --timeout 600 --workers 1 --access-logfile '-' --error-logfile '-' --capture-output wsgi:app
```

Use one worker for production because the app initializes and writes SQLite-backed
settings and canonical-refresh tables during WSGI startup. Multiple workers can race on
those SQLite files during cold start and terminate before the App Service warmup probe
succeeds, causing `ContainerTimeout`, `exit code: 1`, and temporary site blocking. The
`--capture-output` flag sends worker boot exceptions to the App Service log stream.
`PYTHONPATH` includes the workflow-vendored `.python_packages` directory so imports such
as `openpyxl`, `flask`, and `gunicorn` still resolve if the instance starts from a plain
ZIP deployment without an Oryx virtual environment.

## One-time Azure setup

Replace the placeholder values before running:

```powershell
$RESOURCE_GROUP = "jira-reporting-rg"
$PLAN = "jira-reporting-plan"
$APP = "jira-reporting-app"
$LOCATION = "eastus"
```

Create the Linux App Service plan and app:

```powershell
az group create --name $RESOURCE_GROUP --location $LOCATION
az appservice plan create --name $PLAN --resource-group $RESOURCE_GROUP --sku B1 --is-linux
az webapp create --resource-group $RESOURCE_GROUP --plan $PLAN --name $APP --runtime "PYTHON:3.11"
```

Configure the startup command to use the repo script:

```powershell
az webapp config set --resource-group $RESOURCE_GROUP --name $APP --startup-file "startup.txt"
```

Set required app settings:

```powershell
az webapp config appsettings set --resource-group $RESOURCE_GROUP --name $APP --settings REPORT_HTML_DIR=report_html SCM_DO_BUILD_DURING_DEPLOYMENT=true
```

If the app needs Jira-backed refresh features in Azure, also set:

```powershell
az webapp config appsettings set --resource-group $RESOURCE_GROUP --name $APP --settings JIRA_SITE=<your-site> JIRA_EMAIL=<your-email> JIRA_API_TOKEN=<your-token> JIRA_BOARD=<your-board> JIRA_PROJECT_KEYS=<comma-separated-project-keys>
```

## Deploy code

From the repo root:

```powershell
az webapp up --name $APP --resource-group $RESOURCE_GROUP --runtime "PYTHON:3.11" --sku B1
```

If the app already exists and you prefer zip deploy:

```powershell
Compress-Archive -Path * -DestinationPath app.zip -Force
az webapp deploy --resource-group $RESOURCE_GROUP --name $APP --src-path app.zip --type zip
```

## Verify deployment

Open (production example):

```text
https://epreporting.azurewebsites.net/introduction.html
https://epreporting.azurewebsites.net/dashboard.html
https://epreporting.azurewebsites.net/settings/team-capacity-planner
```

The Team Capacity Planner route serves from `REPORT_HTML_DIR` but promotes the
tracked root `team_capacity_planner.html` into that directory if the published
copy is missing or stale. Commit the root HTML source and the server route
change before pushing to `main`; do not rely on an untracked local
`report_html/team_capacity_planner.html` copy for Azure.

Stream logs:

```powershell
az webapp log tail --resource-group $RESOURCE_GROUP --name $APP
```

If the platform log shows `Container has finished running with exit code: 1` during
startup, check the log stream first. The expected healthy boot path is that Gunicorn
continues running after the warmup probe and `/introduction.html` returns HTTP 200. If
the site is temporarily blocked because of repeated cold-start failures, wait for the
reported unblock time, then restart once after the fixed startup command is deployed.

Restart after config changes:

```powershell
az webapp restart --resource-group $RESOURCE_GROUP --name $APP
```

## Modify the persistent SQLite database through SCM/Kudu

Use `azure_scm_sqlite.py` to send a local UTF-8 SQL file to the App Service SCM
endpoint and execute it against a database under persistent `/home`. The script obtains
short-lived publishing credentials through the authenticated Azure CLI; it does not
store credentials in the repository. Temporary `.epr-sqlite-*` runner and SQL files
are removed from `/home/data` after each attempt.

Sign in and preview first (preview is the default and does not execute the SQL):

```powershell
az login
python azure_scm_sqlite.py --resource-group "<resource-group>" --app-name "<app-name>" --sql-file ".\change.sql"
```

Review the database path, SQL SHA-256, statement count, schema-change flag, and
`quick_check_before` result. Then apply the exact same file:

```powershell
python azure_scm_sqlite.py --resource-group "<resource-group>" --app-name "<app-name>" --sql-file ".\change.sql" --apply
```

The default target is `/home/data/assignee_hours_capacity.db`. Override it only with an
absolute persistent path such as `--db-path /home/data/jira_sync_cache.db`. SQL is run
inside `BEGIN IMMEDIATE`; the runner performs SQLite `quick_check` before and after the
statements and rolls back on failure. SQL files cannot include their own transaction
control, `ATTACH`, `DETACH`, `VACUUM`, or `PRAGMA`.

A full database backup is optional because the production capacity database may be
several gigabytes and Azure storage is limited:

```powershell
python azure_scm_sqlite.py --resource-group "<resource-group>" --app-name "<app-name>" --sql-file ".\change.sql" --apply --backup-path "/home/data/backups/assignee_hours_capacity.before-change.db"
```

Only use `--backup-path` after confirming enough free space. Schema-changing SQL is
blocked unless `--allow-schema-changes` is supplied. Before using that flag, complete
the local changelog, authoritative schema snapshot, and production migration steps in
`DATABASE_SCHEMA_MIGRATION_PROTOCOL.md`; this SCM utility does not replace that audit
trail. Stop or quiesce application writers for a long migration to avoid the
60-second SQLite write-lock timeout.

## Notes

- Use Azure App Service for Linux. Microsoft Learn states Python on App Service on Windows is no longer supported.
- The normal local rebuild path currently depends on a populated canonical SQLite run state. Deploying the existing generated `report_html` assets is the safer first release path.
- Keep live mutable EPR data in `/home/data`, not `/home/site/wwwroot`. Set `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db` so deploys do not overwrite the SQLite database. The canonical sync cache also writes SQLite journal files, so set `JIRA_SYNC_DB_PATH=/home/data/jira_sync_cache.db` or allow the server to fall back there when the deployed app root is read-only.
- `report_server.py` trims accidental quotes around `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, creates the parent folder before the first SQLite connection, and verifies the folder is writable. If the configured path is not writable, Azure falls back to `$HOME/data/assignee_hours_capacity.db` and writes a warning to stderr instead of killing the Gunicorn worker during import.
- `report_server.py` applies the same writable-path guard to `JIRA_SYNC_DB_PATH`; if the default `jira_sync_cache.db` under `/home/site/wwwroot` cannot be created because `WEBSITE_RUN_FROM_PACKAGE` is read-only, the sync cache initializes under `$HOME/data/jira_sync_cache.db` instead of failing Gunicorn startup.
- Colossal Refresh stores canonical rows in `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, but generated report artifacts must still be written/synced from the app root. If a refresh completes but reports look stale, verify the app setting, then check that `canonical_refresh_state.last_success_run_id` points at the newest successful run and restart the app after any setting change.
- If Azure restarts while a Colossal Refresh is queued or running, the next `/api/canonical-refresh/current` read or new refresh start marks that orphaned `running` row as `failed` when no in-process worker owns it. The previous successful canonical snapshot remains active, and users can immediately start a new refresh instead of waiting for the manual stuck-run timer.
- If Azure's package mount is read-only during Colossal Refresh, compatibility artifacts such as `1_jira_work_items_export.xlsx`, `2_jira_subtask_worklogs.xlsx`, `3_jira_subtask_worklog_rollup.xlsx`, `nested view.xlsx`, and `jira_exports.db` are written under `$HOME/data/canonical_artifacts` unless `JIRA_CANONICAL_ARTIFACT_DIR` points somewhere else writable.
- Report HTML promotion from root files into `report_html` is best-effort on Azure. If a runtime-promoted `report_html` copy has a newer filesystem timestamp than a freshly deployed root report source, `report_server.py` compares file content before deciding to skip promotion so report-only deploys do not keep serving stale HTML. If `WEBSITE_RUN_FROM_PACKAGE` makes `/home/site/wwwroot` read-only, the server logs a warning and serves the packaged `report_html` copy, the known tracked root report source when no `report_html` copy exists, or a generated dashboard fallback under `$HOME/data/generated_reports`, instead of returning HTTP 500/404 for known reports.
- If you later want Azure to rebuild reports dynamically, validate the writable database path, canonical refresh bootstrap behavior, and generated-report sync path first.
- **WAL mode is disabled everywhere (`db_journal_mode.py`).** Azure mounts `/home` over an SMB/CIFS network share, where SQLite WAL mode (`-shm`/`-wal` shared memory) is unreliable and corrupts the database — the canonical refresh then fails at the *Compare cached modification state* stage with `database disk image is malformed`, requiring a manual `sqlite3 .recover` pass to restore the DB. Rather than run two code paths, `db_journal_mode.apply_journal_mode()` now always selects a rollback journal (`PRAGMA journal_mode=DELETE`) on every host — local and Azure — so no `-wal`/`-shm` sidecar files are ever created and local behavior always matches production. All canonical/sync/exports/migration connections route journal-mode selection through this helper. Override: `EPR_FORCE_WAL=1` forces WAL anywhere (debugging only, not recommended).
- **Recovering an already-corrupt production DB** (the code change only prevents *future* corruption): SSH/Kudu into the instance, run `sqlite3 <db> "PRAGMA integrity_check;"` on `assignee_hours_capacity.db` and `jira_sync_cache.db`. If only `jira_sync_cache.db` is malformed, delete it plus its `-wal`/`-shm` siblings and run a full refresh to rebuild the cache. If `assignee_hours_capacity.db` is malformed, recover with `sqlite3 bad.db ".recover" | sqlite3 recovered.db`, verify `integrity_check`, then swap it in (keep the corrupt copy until verified).

## Business Logic

Azure uses the executable `startup.txt` shell script to run `wsgi:app` with one Gunicorn worker. SQLite-backed runtime state must live under persistent `/home/data`, while generated reports and static assets are served from the deployed application root and `report_html`. The startup script prepends the deploy ZIP's `.python_packages` directory to `PYTHONPATH` so Python packages remain available without relying on an instance-local virtual environment. The workflow preserves executable mode before zipping to avoid `Permission denied` failures when `WEBSITE_RUN_FROM_PACKAGE` mounts `/home/site/wwwroot` read-only.

## Business Cases

The Azure deployment hosts the EPR Tool for production users. Persistent DB storage preserves managed projects, seating, page categories, product releases, and canonical Jira refresh data across code deploys and App Service restarts.

## Examples

With `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db` and `JIRA_SYNC_DB_PATH=/home/data/jira_sync_cache.db`, a deploy updates code under the app root while keeping the 3.3GB production SQLite DB and the canonical sync cache intact under `/home/data`. If either setting is accidentally entered with surrounding quotes, the server strips the quotes before opening SQLite. A successful Colossal Refresh then updates canonical tables in `/home/data` and serves regenerated report HTML from the app root.

## Explanations

Azure extracts the deployed app to a runtime directory and starts Gunicorn. Anything outside `/home` is disposable. `/home/site/wwwroot` may be affected by deployment packaging, so live mutable SQLite files should use `/home/data`; report HTML remains served through the application root and `report_html`.

## Front-end UI Fields

No Azure-only UI fields are added. Relevant production screens are `/settings/canonical-refresh`, report pages under `report_html`, and settings pages that read/write the persistent capacity DB.

- **"Create DB backup before Full Refresh" is disabled.** The `/settings/canonical-refresh` page has a checkbox that used to copy the entire (multi-GB) capacity DB into `backups/canonical_refresh/` before a Full Refresh. This full-database copy filled up disk on production (Azure App Service storage is limited and `/home/data` already holds the live multi-GB DB). The checkbox is now rendered `disabled` and `_start_canonical_refresh_async()` in `report_server.py` forces `should_create_backup = False` regardless of the request payload, so no full-DB backup is ever written by this workflow. The backup helper (`_create_canonical_refresh_db_backup`) and API fields (`db_backup_requested`, `db_backup_created`, `db_backup_path`) remain in the code for compatibility but are always inert.

## Script Files

- `startup.txt` — production shell startup script, Gunicorn command, and `.python_packages` import path.
- `wsgi.py` — Flask app entry point.
- `report_server.py` — server routes, DB path resolution, refresh orchestration, and report serving.
- `db_journal_mode.py` - shared SQLite journal-mode helper imported during WSGI startup.
- `azure_scm_sqlite.py` — local Azure CLI/SCM controller that uploads SQL and invokes the remote runner.
- `azure_scm_sqlite_runner.py` — standard-library-only transactional SQLite runner uploaded temporarily to `/home/data`.
- `.github/workflows/azure-appservice-deploy.yml` — GitHub Actions deployment packaging, dependency vendoring, and executable startup mode preservation.
- `AZURE_APP_SERVICE.md` — Azure operations runbook.

## Dependent & Impacted Files

`generate_assignee_hours_report.py`, canonical refresh generators, report HTML outputs, and SQLite-backed settings modules depend on the Azure DB path being persistent and writable. `requirements.txt` is also a production dependency because the workflow vendors it into `.python_packages` for ZIP startup.

## Table Schema

Azure does not define new tables directly. It hosts SQLite tables from `assignee_hours_capacity.db`, including managed project/settings tables, canonical refresh tables, seating planner tables, page/category tables, and product release tables.

## Data Flow

1. GitHub Actions deploys tracked code to Azure.
2. Azure starts `wsgi:app` through `startup.txt`.
3. Azure runs executable `startup.txt`; the script prepends `.python_packages/lib/site-packages` to `PYTHONPATH` and starts Gunicorn.
4. `report_server.py` resolves, normalizes, and validates `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH` and `JIRA_SYNC_DB_PATH`.
5. Runtime settings and canonical refresh data are read/written in `/home/data/assignee_hours_capacity.db`.
6. Canonical sync-cache state is read/written in `/home/data/jira_sync_cache.db` when the app root is read-only or that path is configured explicitly.
7. Generated reports are synced into the app-served `report_html` path during local builds and deployment packaging; runtime promotion is skipped with a warning when Azure's package mount is read-only, known root report sources remain servable as a fallback, and Dashboard can be regenerated into `$HOME/data/generated_reports/dashboard.html`.
