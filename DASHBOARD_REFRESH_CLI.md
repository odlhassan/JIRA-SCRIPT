# CLI commands to test RMI Dashboard refresh features (local server)

Default server: `http://127.0.0.1:3000` (override with `REPORT_PORT` or `--port` when starting).

## 1. Start the report server

```powershell
cd "D:\JIRA SCRIPT"
python run_server.py
# Or with custom port:
python run_server.py --port 8000
```

Or with env:
```powershell
$env:REPORT_PORT = "8000"
python run_server.py
```

---

## 2. Per-epic refresh (refresh one epic from Jira)

Replace `EPIC_KEY` with a real epic key (e.g. `DIGITALLOG-123`).

**PowerShell:**
```powershell
$base = "http://127.0.0.1:3000"
$epicKey = "EPIC_KEY"
$body = @{ from = "2025-01-01"; to = "2025-12-31" } | ConvertTo-Json
Invoke-RestMethod -Uri "$base/api/dashboard/refresh-epic/$epicKey" -Method POST -Body $body -ContentType "application/json"
```

**curl (Windows cmd or Git Bash):**
```bash
set BASE=http://127.0.0.1:3000
set EPIC_KEY=EPIC_KEY
curl -X POST "%BASE%/api/dashboard/refresh-epic/%EPIC_KEY%" -H "Content-Type: application/json" -d "{\"from\":\"2025-01-01\",\"to\":\"2025-12-31\"}"
```

**curl (PowerShell):**
```powershell
$base = "http://127.0.0.1:3000"
curl.exe -X POST "$base/api/dashboard/refresh-epic/EPIC_KEY" -H "Content-Type: application/json" -d '{\"from\":\"2025-01-01\",\"to\":\"2025-12-31\"}'
```

Response: `{ "ok": true, "epic": {...}, "stories": [...], "subtasks": [...], "bug_subtasks": [...] }`  
This updates `jira_exports.db` (work_items, subtask_worklogs, subtask_worklog_rollup) and returns enriched data.

---

## 3. Full / Smart pipeline refresh

### Start a refresh (Full or Smart)

**Full refresh:**
```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh" -Method POST -ContentType "application/json" -Body '{"mode":"full"}'
# Returns 202 with run_id, e.g. { "ok": true, "run_id": "abc123..." }
```

**Smart refresh (incremental):**
```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh" -Method POST -ContentType "application/json" -Body '{"mode":"smart"}'
```

**curl:**
```bash
curl -X POST "http://127.0.0.1:3000/api/dashboard/refresh" -H "Content-Type: application/json" -d "{\"mode\":\"full\"}"
curl -X POST "http://127.0.0.1:3000/api/dashboard/refresh" -H "Content-Type: application/json" -d "{\"mode\":\"smart\"}"
```

### Poll progress

Replace `RUN_ID` with the `run_id` from the start response.

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh/RUN_ID" -Method GET
```

```bash
curl "http://127.0.0.1:3000/api/dashboard/refresh/RUN_ID"
```

Response includes: `status`, `progress_step`, `progress_pct`, `error_message`. Poll every ~1.2s until `status` is `success`, `failed`, or `canceled`.

### Get current active run

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh/current" -Method GET
```

```bash
curl "http://127.0.0.1:3000/api/dashboard/refresh/current"
```

### Cancel active run

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/cancel" -Method POST
```

```bash
curl -X POST "http://127.0.0.1:3000/api/dashboard/cancel"
```

### Get last run (for Resume / Start Fresh UI)

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh/last" -Method GET
```

```bash
curl "http://127.0.0.1:3000/api/dashboard/refresh/last"
```

### Resume from last failed run

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh" -Method POST -ContentType "application/json" -Body '{"mode":"full","resume":true}'
# Use same "mode" as the failed run (full or smart)
```

### Start from scratch (no resume)

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/dashboard/refresh" -Method POST -ContentType "application/json" -Body '{"mode":"full","resume":false}'
```

---

## 4. Open dashboard in browser

After server is running:

- **RMI Dashboard (report):** http://127.0.0.1:3000/dashboard.html  
- **Root (redirects to dashboard):** http://127.0.0.1:3000/

In the UI you get:
- **Per-epic:** Refresh icon on each epic card → re-fetches that epic from Jira, updates DB, re-renders.
- **Header:** “Refresh” dropdown (Full / Smart), Cancel, progress bar, status text, and Resume / Start Fresh when the last run failed.

---

## 5. Quick one-liner checks (PowerShell)

```powershell
# Current run
(Invoke-RestMethod "http://127.0.0.1:3000/api/dashboard/refresh/current") | ConvertTo-Json

# Last run
(Invoke-RestMethod "http://127.0.0.1:3000/api/dashboard/refresh/last") | ConvertTo-Json
```
