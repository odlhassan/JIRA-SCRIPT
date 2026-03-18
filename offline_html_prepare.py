"""
Prepare Offline HTML: build a timestamped folder of report HTML files with inlined data
and companion JSON files for offline use. Does not modify any generator or existing report logic.
"""
from __future__ import annotations

import json
import re
import shutil
import threading
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

# Default reports for "Prepare Offline HTML" (key = filename stem for API; file = full filename).
DEFAULT_OFFLINE_REPORTS = [
    {"key": "nested_view_report", "file": "nested_view_report.html", "title": "Nested View Report"},
    {"key": "employee_performance_report", "file": "employee_performance_report.html", "title": "Employee Performance"},
    {"key": "rnd_data_story", "file": "rnd_data_story.html", "title": "RnD Data Story"},
    {"key": "planned_rmis_report", "file": "planned_rmis_report.html", "title": "Planned RMIs"},
    {"key": "planned_vs_dispensed_report", "file": "planned_vs_dispensed_report.html", "title": "Approved vs Planned"},
    {"key": "ipp_meeting_dashboard", "file": "ipp_meeting_dashboard.html", "title": "IPP Meeting Dashboard"},
    {"key": "missed_entries", "file": "missed_entries.html", "title": "Missed Entries Report"},
    {"key": "assignee_hours_report", "file": "assignee_hours_report.html", "title": "Assignee Hours Report"},
    {"key": "rlt_leave_report", "file": "rlt_leave_report.html", "title": "RLT Leave Report"},
    {"key": "leaves_planned_calendar", "file": "leaves_planned_calendar.html", "title": "Leaves Planned Calendar"},
    {"key": "original_estimates_hierarchy_report", "file": "original_estimates_hierarchy_report.html", "title": "Epic Estimate Report"},
]

# Reports that have embedded reportData in HTML (read + optional date filter).
EMBEDDED_REPORTS = {
    "nested_view_report.html",
    "planned_rmis_report.html",
    "missed_entries.html",
    "rlt_leave_report.html",
    "leaves_planned_calendar.html",
    "gantt_chart_report.html",
    "phase_rmi_gantt_report.html",
}

# Reports that load data via APIs (we call APIs and inject fetch override).
API_REPORTS = {
    "employee_performance_report.html",
    "rnd_data_story.html",
    "planned_vs_dispensed_report.html",
    "assignee_hours_report.html",
    "original_estimates_hierarchy_report.html",
    "ipp_meeting_dashboard.html",
}


def get_default_report_keys() -> list[str]:
    """Return the default list of report keys (filename stems) for the UI."""
    return [r["key"] for r in DEFAULT_OFFLINE_REPORTS]


def get_default_reports_for_ui() -> list[dict[str, str]]:
    """Return default reports as list of { key, file, title } for modal checkboxes."""
    return list(DEFAULT_OFFLINE_REPORTS)


def _folder_name_from_now() -> str:
    # Use hyphen in time so the path is valid on Windows (colons not allowed in dir names).
    return datetime.now(timezone.utc).strftime("%d %b %Y %H-%M-%S")


def _normalize_report_key_to_filename(key: str) -> str:
    """Convert API key to filename (e.g. nested_view_report -> nested_view_report.html)."""
    key = (key or "").strip()
    if not key:
        return ""
    if key.endswith(".html"):
        return key
    return key + ".html"


def _extract_json_object_after(html: str, prefix: str) -> dict[str, Any] | None:
    """Find prefix in html, then parse a balanced { ... } and return as dict."""
    idx = html.find(prefix)
    if idx == -1:
        return None
    start = idx + len(prefix)
    if start >= len(html) or html[start] != "{":
        return None
    depth = 0
    i = start
    in_string = None
    escape = False
    while i < len(html):
        c = html[i]
        if escape:
            escape = False
            i += 1
            continue
        if c == "\\" and in_string:
            escape = True
            i += 1
            continue
        if in_string:
            if c == in_string:
                in_string = None
            i += 1
            continue
        if c in ('"', "'"):
            in_string = c
            i += 1
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(html[start : i + 1])
                except json.JSONDecodeError:
                    return None
        i += 1
    return None


def _extract_report_data_from_html(html: str) -> dict[str, Any] | None:
    """Extract reportData (or similar) from HTML. Returns None if not found."""
    for prefix in ("const reportData = ", "window.REPORT_DATA = "):
        out = _extract_json_object_after(html, prefix)
        if out is not None:
            return out
    return None


def _filter_rows_by_date(
    payload: dict[str, Any],
    from_date: str,
    to_date: str,
    row_date_start_key: str = "jira_start_date",
    row_date_end_key: str = "jira_due_date",
) -> dict[str, Any]:
    """Filter payload rows by date range. Supports rows with start/end date fields. Returns new dict."""
    from_d = from_date[:10] if from_date else ""
    to_d = to_date[:10] if to_date else ""
    if not from_d or not to_d:
        return payload

    out = dict(payload)
    rows = out.get("rows")
    if not isinstance(rows, list):
        return payload
    filtered = []
    for r in rows:
        if not isinstance(r, dict):
            filtered.append(r)
            continue
        start_val = (r.get(row_date_start_key) or "")[:10]
        end_val = (r.get(row_date_end_key) or "")[:10]
        # Include if any overlap: row start <= to_d and row end >= from_d (or empty dates = include)
        if not start_val and not end_val:
            filtered.append(r)
            continue
        if start_val and start_val > to_d:
            continue
        if end_val and end_val < from_d:
            continue
        filtered.append(r)
    out["rows"] = filtered
    return out


def _filter_items_by_date(
    payload: dict[str, Any],
    from_date: str,
    to_date: str,
) -> dict[str, Any]:
    """Filter payload 'items' by planned_start/planned_end. Returns new dict."""
    from_d = from_date[:10] if from_date else ""
    to_d = to_date[:10] if to_date else ""
    if not from_d or not to_d:
        return payload
    out = dict(payload)
    items = out.get("items")
    if not isinstance(items, list):
        return payload
    filtered = []
    for r in items:
        if not isinstance(r, dict):
            filtered.append(r)
            continue
        start_val = (r.get("planned_start") or "")[:10]
        end_val = (r.get("planned_end") or "")[:10]
        if not start_val and not end_val:
            filtered.append(r)
            continue
        if start_val and start_val > to_d:
            continue
        if end_val and end_val < from_d:
            continue
        filtered.append(r)
    out["items"] = filtered
    return out


def _inject_embedded_data_into_html(html: str, new_data: dict[str, Any]) -> str:
    """Replace existing reportData assignment in HTML with new_data (JSON-inlined)."""
    json_str = json.dumps(new_data, ensure_ascii=False)
    # Replace const reportData = { ... }; with const reportData = <new>; (allow }; followed by newline, </ or end)
    pattern = re.compile(
        r"const\s+reportData\s*=\s*\{[\s\S]*?\};\s*(?=\n|</|$)",
        re.MULTILINE,
    )
    replacement = f"const reportData = {json_str};"
    if pattern.search(html):
        return pattern.sub(replacement, html, count=1)
    # Prepend before first <script> so report has data
    first_script = html.find("<script>")
    if first_script != -1:
        return html[:first_script] + f"<script>\n{replacement}\n</script>\n" + html[first_script:]
    return html


def _make_fetch_override_script(
    bundle: dict[str, Any],
    from_date: str = "",
    to_date: str = "",
) -> str:
    """Return a script that sets OFFLINE_API_BUNDLE and overrides fetch to serve from it.
    When from_date/to_date are set, seeds date inputs (#from, #to) so the report uses the same range as the bundle."""
    bundle_json = json.dumps(bundle, ensure_ascii=False)
    bundle_json_escaped = bundle_json.replace("</", "<\\/")
    from_esc = json.dumps(from_date[:10] if from_date else "")
    to_esc = json.dumps(to_date[:10] if to_date else "")
    date_seed = ""
    if from_date and to_date:
        date_seed = f"""
  window.OFFLINE_DATE_FROM = {from_esc};
  window.OFFLINE_DATE_TO = {to_esc};
  function applyOfflineDates() {{
    if (window.OFFLINE_DATE_FROM && window.OFFLINE_DATE_TO) {{
      var fromEl = document.getElementById('from') || document.getElementById('date-filter-from') || document.getElementById('from-date');
      var toEl = document.getElementById('to') || document.getElementById('date-filter-to') || document.getElementById('to-date');
      if (fromEl) fromEl.value = window.OFFLINE_DATE_FROM;
      if (toEl) toEl.value = window.OFFLINE_DATE_TO;
    }}
  }}
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', applyOfflineDates);
  else applyOfflineDates();
"""
    return f"""<script>
(function() {{{date_seed}
  window.OFFLINE_API_BUNDLE = {bundle_json_escaped};
  var _fetch = window.fetch;
  window.fetch = function(url, opts) {{
    var u = (url || '').toString();
    var method = (opts && opts.method) || 'GET';
    var full = method + ' ' + u;
    var pathOnly = method + ' ' + u.split('?')[0];
    var data = window.OFFLINE_API_BUNDLE[full] || window.OFFLINE_API_BUNDLE[pathOnly];
    if (data !== undefined) {{
      return Promise.resolve({{ ok: true, json: function() {{ return Promise.resolve(data); }}, text: function() {{ return Promise.resolve(JSON.stringify(data)); }} }});
    }}
    return _fetch.apply(this, arguments);
  }};
}})();
</script>"""


def _inject_fetch_override_into_html(
    html: str,
    bundle: dict[str, Any],
    from_date: str = "",
    to_date: str = "",
) -> str:
    """Inject the fetch-override script at the start of <body>."""
    script = _make_fetch_override_script(bundle, from_date=from_date, to_date=to_date)
    body_start = html.find("<body>")
    if body_start != -1:
        insert = body_start + len("<body>")
        return html[:insert] + "\n" + script + "\n" + html[insert:]
    head_end = html.find("</head>")
    if head_end != -1:
        return html[:head_end] + "\n" + script + "\n</head>" + html[head_end:]
    return script + "\n" + html


def run_prepare_job(
    base_dir: Path,
    from_date: str,
    to_date: str,
    report_keys: list[str],
    resolve_sources: Callable[[Path], dict[str, Path]],
    sync_assets: Callable[[Path, Path], None],
    fetch_api: Callable[[str, str, dict[str, str] | None], Any],
    update_progress: Callable[[int, int, str], None] | None = None,
) -> tuple[Path, list[str]]:
    """
    Run the prepare-offline job. Returns (output_folder_path, list of error messages).
    Does not raise; errors are collected and returned.
    """
    folder_name = _folder_name_from_now()
    output_dir = base_dir / "offline_bundles" / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    sources = resolve_sources(base_dir)
    total = len(report_keys)
    done = 0

    for key in report_keys:
        filename = _normalize_report_key_to_filename(key)
        if not filename or filename not in sources:
            errors.append(f"Unknown or missing report: {key}")
            if update_progress:
                update_progress(done, total, f"Skipped {key}")
            continue
        source_path = sources[filename]
        if not source_path.exists() or not source_path.is_file():
            errors.append(f"Source file not found: {source_path}")
            if update_progress:
                update_progress(done, total, f"Skip {filename}")
            continue
        try:
            html_content = source_path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            errors.append(f"Read {filename}: {e}")
            if update_progress:
                update_progress(done, total, f"Error {filename}")
            done += 1
            continue

        if filename in EMBEDDED_REPORTS:
            payload = _extract_report_data_from_html(html_content)
            if payload is None:
                payload = {}
            if filename == "missed_entries.html":
                payload = _filter_rows_by_date(payload, from_date, to_date)
            elif filename in ("phase_rmi_gantt_report.html",):
                payload = _filter_items_by_date(payload, from_date, to_date)
            elif "rows" in payload and payload["rows"] and isinstance(payload["rows"][0], dict):
                row0 = payload["rows"][0]
                if "planned_start" in row0 or "jira_start_date" in row0:
                    payload = _filter_rows_by_date(
                        payload, from_date, to_date,
                        row_date_start_key="planned_start" if "planned_start" in row0 else "jira_start_date",
                        row_date_end_key="planned_end" if "planned_end" in row0 else "jira_due_date",
                    )
            html_content = _inject_embedded_data_into_html(html_content, payload)
            (output_dir / filename).write_text(html_content, encoding="utf-8")
            (output_dir / filename.replace(".html", ".json")).write_text(
                json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
            )
        elif filename in API_REPORTS:
            bundle = _fetch_api_bundle_for_report(
                filename, from_date, to_date, fetch_api
            )
            html_content = _inject_fetch_override_into_html(
                html_content, bundle, from_date=from_date, to_date=to_date
            )
            (output_dir / filename).write_text(html_content, encoding="utf-8")
            (output_dir / filename.replace(".html", ".json")).write_text(
                json.dumps(bundle, indent=2, ensure_ascii=False), encoding="utf-8"
            )
        else:
            # Copy as-is and try to extract + write JSON for audit
            payload = _extract_report_data_from_html(html_content)
            if payload is not None:
                (output_dir / filename.replace(".html", ".json")).write_text(
                    json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
                )
            (output_dir / filename).write_text(html_content, encoding="utf-8")

        done += 1
        if update_progress:
            update_progress(done, total, filename)

    sync_assets(base_dir, output_dir)
    return output_dir, errors


def _fetch_api_bundle_for_report(
    filename: str,
    from_date: str,
    to_date: str,
    fetch_api: Callable[[str, str, dict[str, str] | None], Any],
) -> dict[str, Any]:
    """Build OFFLINE_API_BUNDLE for an API-driven report. Keys: 'GET /path?query' and 'GET /path' for fallback."""
    bundle: dict[str, Any] = {}
    from_d = (from_date or "")[:10]
    to_d = (to_date or "")[:10]

    def get(path: str, params: dict[str, str] | None = None) -> None:
        q = "&".join(f"{k}={v}" for k, v in sorted((params or {}).items()))
        url = path + ("?" + q if q else "")
        key_full = "GET " + url
        key_path = "GET " + path
        try:
            data = fetch_api("GET", path, params)
            if data is not None:
                bundle[key_full] = data
                bundle[key_path] = data  # fallback when query string order differs
        except Exception:
            bundle[key_full] = {"ok": False, "error": "Failed to fetch"}
            bundle[key_path] = bundle[key_full]

    if filename == "employee_performance_report.html":
        get("/api/performance/settings")
        # Use same project scope as live report (managed projects) so Planned/Actual stats match.
        project_keys: list[str] = []
        try:
            projects_data = fetch_api("GET", "/api/projects", None)
            if isinstance(projects_data, dict) and "projects" in projects_data:
                for p in projects_data.get("projects") or []:
                    if isinstance(p, dict):
                        key = (p.get("project_key") or "").strip().upper()
                        if key:
                            project_keys.append(key)
        except Exception:
            pass
        # Use log_date so actuals are only worklogs in the selected date range (offline = that date only).
        scoped_params: dict[str, str] = {"from": from_d, "to": to_d, "mode": "log_date"}
        if project_keys:
            scoped_params["projects"] = ",".join(sorted(set(project_keys)))
        get("/api/scoped-subtasks", scoped_params)
    elif filename == "rnd_data_story.html":
        get("/api/scoped-subtasks", {"from": from_d, "to": to_d, "mode": "planned_dates"})
        get("/api/capacity", {"from": from_d, "to": to_d})
    elif filename == "planned_vs_dispensed_report.html":
        get("/api/approved-vs-planned-hours/ui-settings")
        get("/api/approved-vs-planned-hours/summary", {"from": from_d, "to": to_d})
        get("/api/approved-vs-planned-hours/details", {"from": from_d, "to": to_d})
    elif filename == "assignee_hours_report.html":
        get("/api/capacity", {"from": from_d, "to": to_d})
        get("/api/capacity/profiles")
        get("/api/actual-hours/aggregate", {"from": from_d, "to": to_d})
    elif filename == "original_estimates_hierarchy_report.html":
        get("/api/original-estimates/filter-options")
        get("/api/original-estimates/summary", {"from": from_d, "to": to_d})
    elif filename == "ipp_meeting_dashboard.html":
        get("/api/ipp-meeting-dashboard/data")
    else:
        pass
    return bundle


def create_zip_of_folder(folder_path: Path) -> Path:
    """Create a zip file of the folder. Returns path to the zip. Caller can stream or delete folder after."""
    zip_path = folder_path.parent / f"{folder_path.name}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in folder_path.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(folder_path))
    return zip_path
