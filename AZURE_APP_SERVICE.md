# Azure App Service Deployment

This project can run on Azure App Service for Linux as a Python 3.11 web app.

## GitHub Actions → production

- **Live URL:** `https://epreporting.azurewebsites.net/`
- Workflow: `.github/workflows/azure-appservice-deploy.yml` (runs on push to `main` or `master`, and manual `workflow_dispatch`).
- The deploy ZIP is built from the **GitHub checkout** after staging (excludes tests, `node_modules`, `handover`, `offline_bundles`, root `backup/`, etc.). **Commit and push every file the running app depends on** — anything not in Git will not ship. Do not rely on unpushed local builds.
- The workflow runs `pip install -r requirements.txt --target staging/.python_packages/lib/site-packages` before zipping. This makes production dependency imports independent of whether Azure creates `/home/site/wwwroot/antenv` during deployment. It also marks `startup.txt` executable before creating the ZIP so App Service can launch it from a read-only `WEBSITE_RUN_FROM_PACKAGE` mount.
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

## Notes

- Use Azure App Service for Linux. Microsoft Learn states Python on App Service on Windows is no longer supported.
- The normal local rebuild path currently depends on a populated canonical SQLite run state. Deploying the existing generated `report_html` assets is the safer first release path.
- Keep live mutable EPR data in `/home/data`, not `/home/site/wwwroot`. Set `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db` so deploys do not overwrite the SQLite database.
- `report_server.py` trims accidental quotes around `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, creates the parent folder before the first SQLite connection, and verifies the folder is writable. If the configured path is not writable, Azure falls back to `$HOME/data/assignee_hours_capacity.db` and writes a warning to stderr instead of killing the Gunicorn worker during import.
- Colossal Refresh stores canonical rows in `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`, but generated report artifacts must still be written/synced from the app root. If a refresh completes but reports look stale, verify the app setting, then check that `canonical_refresh_state.last_success_run_id` points at the newest successful run and restart the app after any setting change.
- If you later want Azure to rebuild reports dynamically, validate the writable database path, canonical refresh bootstrap behavior, and generated-report sync path first.

## Business Logic

Azure uses the executable `startup.txt` shell script to run `wsgi:app` with one Gunicorn worker. SQLite-backed runtime state must live under persistent `/home/data`, while generated reports and static assets are served from the deployed application root and `report_html`. The startup script prepends the deploy ZIP's `.python_packages` directory to `PYTHONPATH` so Python packages remain available without relying on an instance-local virtual environment. The workflow preserves executable mode before zipping to avoid `Permission denied` failures when `WEBSITE_RUN_FROM_PACKAGE` mounts `/home/site/wwwroot` read-only.

## Business Cases

The Azure deployment hosts the EPR Tool for production users. Persistent DB storage preserves managed projects, seating, page categories, product releases, and canonical Jira refresh data across code deploys and App Service restarts.

## Examples

With `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH=/home/data/assignee_hours_capacity.db`, a deploy updates code under the app root while keeping the 3.3GB production SQLite DB intact under `/home/data`. If the setting is accidentally entered as `"/home/data/assignee_hours_capacity.db"`, the server strips the quotes before opening SQLite. A successful Colossal Refresh then updates canonical tables in `/home/data` and serves regenerated report HTML from the app root.

## Explanations

Azure extracts the deployed app to a runtime directory and starts Gunicorn. Anything outside `/home` is disposable. `/home/site/wwwroot` may be affected by deployment packaging, so live mutable SQLite files should use `/home/data`; report HTML remains served through the application root and `report_html`.

## Front-end UI Fields

No Azure-only UI fields are added. Relevant production screens are `/settings/canonical-refresh`, report pages under `report_html`, and settings pages that read/write the persistent capacity DB.

## Script Files

- `startup.txt` — production shell startup script, Gunicorn command, and `.python_packages` import path.
- `wsgi.py` — Flask app entry point.
- `report_server.py` — server routes, DB path resolution, refresh orchestration, and report serving.
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
4. `report_server.py` resolves, normalizes, and validates `JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH`.
5. Runtime settings and canonical refresh data are read/written in `/home/data/assignee_hours_capacity.db`.
6. Generated reports are synced into the app-served `report_html` path.
