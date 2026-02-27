# Runbook

## 1) Install dependency
```powershell
pip install -r requirements_missed_entries.txt
```

## 2) Generate report (default paths)
```powershell
python generate_missed_entries_html.py
```

## 3) Generate report (custom paths)
```powershell
$env:JIRA_EXPORT_XLSX_PATH = "E:\path\to\1_jira_work_items_export.xlsx"
$env:JIRA_MISSED_ENTRIES_HTML_PATH = "E:\path\to\missed_entries.html"
python generate_missed_entries_html.py
```

## 4) Deploy HTML to served folder (example)
```powershell
Copy-Item missed_entries.html report_html\missed_entries.html -Force
```

## 5) Browser cache refresh
Use `Ctrl+F5` after deploying a new report.

## Troubleshooting
- Error: workbook not found
  - Check `JIRA_EXPORT_XLSX_PATH` and file exists.
- UI Excel export button does nothing
  - Ensure browser can load JSDelivr `xlsx` script.
- Jira links look wrong
  - Set `JIRA_SITE` correctly.
