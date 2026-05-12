# Azure App Service Deployment

This project can run on Azure App Service for Linux as a Python 3.11 web app.

## GitHub Actions → production

- **Live URL:** `https://epreporting.azurewebsites.net/`
- Workflow: `.github/workflows/azure-appservice-deploy.yml` (runs on push to `main` or `master`, and manual `workflow_dispatch`).
- The deploy ZIP is built from the **GitHub checkout** after staging (excludes tests, `node_modules`, `handover`, `offline_bundles`, root `backup/`, etc.). **Commit and push every file the running app depends on** — anything not in Git will not ship. Do not rely on unpushed local builds.
- **Backups:** do not commit local backup trees or `*.backup` files (see `.gitignore`). Do not expect backup-only branches to drive production unless you change the workflow deliberately.

## Why a startup command is required

Azure App Service auto-detects Flask only when the app entrypoint is `app.py` or `application.py` with an `app` callable.

This repo exposes the Flask app from `wsgi.py`, so configure Azure to use the included `startup.txt` file.

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

Configure the startup command to use the repo file:

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
```

Stream logs:

```powershell
az webapp log tail --resource-group $RESOURCE_GROUP --name $APP
```

Restart after config changes:

```powershell
az webapp restart --resource-group $RESOURCE_GROUP --name $APP
```

## Notes

- Use Azure App Service for Linux. Microsoft Learn states Python on App Service on Windows is no longer supported.
- The normal local rebuild path currently depends on a populated canonical SQLite run state. Deploying the existing generated `report_html` assets is the safer first release path.
- If you later want Azure to rebuild reports dynamically, validate the writable database path and canonical refresh bootstrap behavior first.
