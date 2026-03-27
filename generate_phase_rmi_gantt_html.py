"""
Generate a team-owner RMI gantt chart HTML report from SQLite snapshot tables.
"""
from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_OUTPUT_HTML = "phase_rmi_gantt_report.html"
DEFAULT_CAPACITY_DB = "assignee_hours_capacity.db"


def _resolve_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return base_dir / path


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    try:
        return round(float(value or 0.0), 2)
    except (TypeError, ValueError):
        return 0.0


def _load_team_rmi_payload(db_path: Path) -> dict[str, Any]:
    if not db_path.exists():
        return {
            "team_names": [],
            "items": [],
            "snapshot_meta": {},
            "source_file": str(db_path),
        }
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        tables = {
            _to_text(r[0]).lower()
            for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }
        if "team_rmi_gantt_items" not in tables:
            return {
                "team_names": [],
                "items": [],
                "snapshot_meta": {},
                "source_file": str(db_path),
            }

        columns = {
            _to_text(r[1]).lower()
            for r in conn.execute("PRAGMA table_info(team_rmi_gantt_items)").fetchall()
        }
        has_epic_status = "epic_status" in columns
        select_sql = """
            SELECT team_name, epic_key, epic_name, epic_url, {epic_status_sql} AS epic_status, project_key,
                   planned_start, planned_end, planned_hours, planned_man_days, story_count, is_unmapped_team, snapshot_utc
            FROM team_rmi_gantt_items
            ORDER BY lower(team_name), lower(epic_name), lower(epic_key)
        """.format(epic_status_sql="epic_status" if has_epic_status else "''")
        rows = conn.execute(select_sql).fetchall()
        items = []
        team_names: list[str] = []
        seen_teams: set[str] = set()
        for row in rows:
            team_name = _to_text(row["team_name"])
            if team_name and team_name not in seen_teams:
                seen_teams.add(team_name)
                team_names.append(team_name)
            items.append(
                {
                    "team_name": team_name,
                    "epic_key": _to_text(row["epic_key"]),
                    "epic_name": _to_text(row["epic_name"]),
                    "epic_url": _to_text(row["epic_url"]),
                    "epic_status": _to_text(row["epic_status"]),
                    "project_key": _to_text(row["project_key"]),
                    "planned_start": _to_text(row["planned_start"]),
                    "planned_end": _to_text(row["planned_end"]),
                    "planned_hours": _to_float(row["planned_hours"]),
                    "planned_man_days": _to_float(row["planned_man_days"]),
                    "story_count": int(row["story_count"] or 0),
                    "is_unmapped_team": int(row["is_unmapped_team"] or 0),
                    "snapshot_utc": _to_text(row["snapshot_utc"]),
                }
            )

        snapshot_meta = {}
        if "team_rmi_gantt_snapshot_meta" in tables:
            meta_row = conn.execute(
                """
                SELECT snapshot_utc, source_work_items_path, total_story_rows, included_story_rows,
                       excluded_missing_epic, excluded_missing_dates, excluded_missing_estimate
                FROM team_rmi_gantt_snapshot_meta
                WHERE id = 1
                """
            ).fetchone()
            if meta_row:
                snapshot_meta = {
                    "snapshot_utc": _to_text(meta_row["snapshot_utc"]),
                    "source_work_items_path": _to_text(meta_row["source_work_items_path"]),
                    "total_story_rows": int(meta_row["total_story_rows"] or 0),
                    "included_story_rows": int(meta_row["included_story_rows"] or 0),
                    "excluded_missing_epic": int(meta_row["excluded_missing_epic"] or 0),
                    "excluded_missing_dates": int(meta_row["excluded_missing_dates"] or 0),
                    "excluded_missing_estimate": int(meta_row["excluded_missing_estimate"] or 0),
                }

        source_file = str(db_path)
        return {
            "team_names": team_names,
            "items": items,
            "snapshot_meta": snapshot_meta,
            "source_file": source_file,
        }
    finally:
        conn.close()


def _build_html(data: dict[str, Any]) -> str:
    payload = json.dumps(data, ensure_ascii=True)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Team Owner RMI Gantt</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,500,0,0">
  <style>
    :root {{
      --bg: #eef3f7;
      --panel: #ffffff;
      --line: #d6e1ea;
      --text: #1f2937;
      --muted: #5f6f7f;
      --title: #0c4054;
      --head: #0f4c5c;
      --head-text: #ffffff;
      --team-col: 280px;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      color: var(--text);
      font-family: "Segoe UI", Tahoma, Verdana, sans-serif;
      background:
        radial-gradient(1000px 320px at 8% -10%, #dbeef7 0%, transparent 62%),
        linear-gradient(180deg, #f4f8fb, var(--bg));
    }}
    .page {{
      max-width: 1880px;
      margin: 0 auto;
      padding: 16px;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 14px 16px;
      margin-bottom: 12px;
    }}
    .title {{
      margin: 0 0 6px;
      font-size: 1.25rem;
      color: var(--title);
    }}
    .meta {{
      margin: 0;
      color: var(--muted);
      font-size: 0.9rem;
    }}
    .controls {{
      margin-top: 10px;
      display: flex;
      flex-wrap: wrap;
      align-items: end;
      gap: 10px;
    }}
    .control {{
      display: flex;
      flex-direction: column;
      gap: 4px;
      min-width: 180px;
    }}
    .control.compact {{
      min-width: 44px;
      align-self: end;
    }}
    .control label {{
      font-size: 0.8rem;
      color: #314756;
      font-weight: 600;
    }}
    .control input,
    .control select {{
      border: 1px solid #b9cad5;
      border-radius: 8px;
      padding: 7px 9px;
      font-size: 0.9rem;
      color: #102c3b;
      background: #fff;
    }}
    .control input:focus,
    .control select:focus {{
      outline: none;
      border-color: #2a6274;
      box-shadow: 0 0 0 2px rgba(42, 98, 116, 0.16);
    }}
    .nav-btn {{
      border: 1px solid #b9cad5;
      border-radius: 8px;
      background: #fff;
      color: #1f4658;
      padding: 0;
      width: 40px;
      height: 39px;
      font-size: 1rem;
      line-height: 1;
      cursor: pointer;
    }}
    .nav-btn:hover {{
      background: #f5f9fc;
    }}
    .nav-btn:focus {{
      outline: none;
      border-color: #2a6274;
      box-shadow: 0 0 0 2px rgba(42, 98, 116, 0.16);
    }}
    .btn {{
      border: 1px solid #255f73;
      background: #0f4c5c;
      color: #fff;
      border-radius: 8px;
      padding: 7px 12px;
      font-size: 0.86rem;
      cursor: pointer;
      height: 34px;
    }}
    .btn.alt {{
      background: #fff;
      color: #255f73;
    }}
    .legend {{
      margin-top: 10px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      color: #506575;
      font-size: 0.78rem;
    }}
    .legend-pill {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border: 1px solid #d8e4ee;
      border-radius: 999px;
      padding: 3px 8px;
      background: #f8fbff;
    }}
    .quick-filters {{
      margin-top: 10px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }}
    .quick-filter-btn {{
      border: 1px solid #c3d4df;
      background: #f8fbff;
      color: #1f4658;
      border-radius: 999px;
      padding: 6px 10px;
      font-size: 0.78rem;
      font-weight: 600;
      cursor: pointer;
      line-height: 1;
    }}
    .quick-filter-btn:hover {{
      background: #eef6fb;
    }}
    .quick-filter-btn.active {{
      border-color: #255f73;
      background: #0f4c5c;
      color: #fff;
    }}
    .chip-dot {{
      width: 10px;
      height: 10px;
      border-radius: 999px;
      background: #93c5fd;
    }}
    .gantt-wrap {{
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 12px;
      overflow: auto;
      max-height: 76vh;
    }}
    .gantt-root {{
      min-width: 860px;
    }}
    .grid {{
      min-width: max-content;
      border-collapse: separate;
      border-spacing: 0;
    }}
    .grid-row {{
      display: grid;
      grid-template-columns: var(--team-col) 1fr;
      min-width: max-content;
    }}
    .cell {{
      border-top: 1px solid var(--line);
      background: #fff;
    }}
    .head .cell {{
      position: sticky;
      top: 0;
      z-index: 40;
      border-top: 0;
      background: var(--head);
      color: var(--head-text);
      font-weight: 700;
    }}
    .head .team-cell {{
      padding: 10px 12px;
      border-right: 1px solid rgba(255,255,255,0.18);
    }}
    .head .timeline-cell {{
      padding: 0;
      z-index: 38;
    }}
    .team-cell {{
      position: sticky;
      left: 0;
      z-index: 22;
      border-right: 1px solid var(--line);
      padding: 10px 12px;
      background: #fff;
      min-width: var(--team-col);
      max-width: var(--team-col);
    }}
    .team-title {{
      font-size: 0.88rem;
      font-weight: 700;
      color: #0f3040;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .team-sub {{
      margin-top: 3px;
      font-size: 0.74rem;
      color: #637a8a;
    }}
    .load-row {{
      margin-top: 7px;
      display: grid;
      grid-auto-flow: column;
      grid-auto-columns: 12px;
      gap: 3px;
      justify-content: start;
      align-items: center;
      overflow: hidden;
      max-width: 252px;
      min-height: 12px;
    }}
    .load-chip {{
      width: 12px;
      height: 10px;
      border-radius: 3px;
      border: 1px solid #d5e0e8;
      background: #eff4f8;
    }}
    .timeline-cell {{
      position: relative;
      overflow: hidden;
      min-height: 48px;
      background: #fff;
    }}
    .lane-timeline {{
      position: relative;
      padding: 8px 8px 10px;
    }}
    .timeline-header {{
      position: relative;
      min-height: 58px;
      background: #f8fbff;
      border-bottom: 1px solid #d4e0e9;
    }}
    .month-block {{
      position: absolute;
      top: 0;
      height: 22px;
      border-right: 1px solid #d0dce5;
      background: rgba(210, 228, 239, 0.35);
      color: #334f60;
      padding: 3px 6px;
      font-size: 0.72rem;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .week-block {{
      position: absolute;
      top: 24px;
      height: 32px;
      border-right: 1px solid #d8e3eb;
      padding: 4px 6px;
      font-size: 0.72rem;
      color: #3b5565;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      background: rgba(255,255,255,0.62);
    }}
    .week-block.current-week {{
      background: rgba(251, 191, 36, 0.28);
      border: 1px solid rgba(180, 83, 9, 0.45);
      color: #7c2d12;
      font-weight: 700;
      z-index: 2;
    }}
    .grid-line {{
      position: absolute;
      top: 0;
      bottom: 0;
      width: 1px;
      background: #e8eef3;
      pointer-events: none;
    }}
    .week-line {{
      position: absolute;
      top: 0;
      bottom: 0;
      width: 1px;
      background: #d8e2ea;
      pointer-events: none;
    }}
    .current-week-band {{
      position: absolute;
      top: 0;
      bottom: 0;
      background: rgba(251, 191, 36, 0.13);
      border-left: 1px solid rgba(180, 83, 9, 0.4);
      border-right: 1px solid rgba(180, 83, 9, 0.4);
      pointer-events: none;
      z-index: 1;
    }}
    .lane-empty {{
      padding: 10px;
      font-size: 0.8rem;
      color: #7b8c99;
    }}
    .card {{
      position: absolute;
      min-height: 66px;
      border-radius: 9px;
      border: 1px solid var(--card-border, rgba(15, 76, 92, 0.32));
      background: linear-gradient(180deg, var(--card-bg-top, rgba(219, 238, 247, 0.82)), var(--card-bg-bottom, rgba(233, 246, 253, 0.95)));
      box-shadow: inset 4px 0 0 var(--card-accent, #7cb6d1), 0 2px 8px rgba(16, 43, 58, 0.08);
      padding: 0;
      overflow: hidden;
    }}
    .card-link {{
      display: block;
      color: inherit;
      text-decoration: none;
      padding: 6px 7px;
      width: 100%;
      height: 100%;
    }}
    .card-link:hover {{
      background: rgba(191, 219, 254, 0.26);
    }}
    .card-title {{
      font-size: 0.76rem;
      color: #103547;
      font-weight: 700;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      margin-bottom: 4px;
    }}
    .card-meta {{
      font-size: 0.68rem;
      color: #375567;
      line-height: 1.32;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .status-pill {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      margin-bottom: 5px;
      padding: 2px 7px;
      border-radius: 999px;
      border: 1px solid var(--pill-border, #bfd3df);
      background: var(--pill-bg, #edf5fa);
      color: var(--pill-text, #1f4658);
      font-size: 0.65rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.03em;
      max-width: 100%;
    }}
    .empty {{
      padding: 16px;
      color: #607282;
      font-size: 0.9rem;
    }}
  </style>
  <link rel="stylesheet" href="shared-nav.css">
</head>
<body>
  <div class="page">
    <section class="panel">
      <h1 class="title">Team Owner RMI Gantt</h1>
      <p class="meta">Generated: <span id="generated-at"></span> | Source: <span id="source-file"></span> | Visible RMIs: <span id="visible-count"></span></p>
      <p class="meta">Included stories: <span id="included-stories"></span> / <span id="total-stories"></span> | Excluded: missing epic <span id="excluded-epic"></span>, missing dates <span id="excluded-dates"></span>, missing estimate <span id="excluded-estimate"></span></p>
      <div class="controls">
        <div class="control">
          <label for="team-select">Team</label>
          <select id="team-select"></select>
        </div>
        <div class="control compact">
          <label for="shift-range-back">Shift</label>
          <button class="nav-btn" type="button" id="shift-range-back" aria-label="Shift current date range to previous month">&#9664;</button>
        </div>
        <div class="control">
          <label for="from-date">From</label>
          <input type="date" id="from-date">
        </div>
        <div class="control">
          <label for="to-date">To</label>
          <input type="date" id="to-date">
        </div>
        <div class="control compact">
          <label for="shift-range-forward">Shift</label>
          <button class="nav-btn" type="button" id="shift-range-forward" aria-label="Shift current date range to next month">&#9654;</button>
        </div>
        <button class="btn alt" type="button" id="apply-range">Apply</button>
        <button class="btn" type="button" id="reset-range">Reset</button>
      </div>
      <div class="quick-filters" id="quick-filters">
        <button class="quick-filter-btn" type="button" id="shortcut-this-year" data-range-preset="this-year">This Year</button>
        <button class="quick-filter-btn" type="button" id="shortcut-this-month" data-range-preset="this-month">This Month</button>
        <button class="quick-filter-btn" type="button" id="shortcut-previous-month" data-range-preset="previous-month">Previous Month</button>
        <button class="quick-filter-btn" type="button" id="shortcut-this-quarter" data-range-preset="this-quarter">This Quarter</button>
        <button class="quick-filter-btn" type="button" id="shortcut-this-week" data-range-preset="this-week">This Week</button>
        <button class="quick-filter-btn" type="button" id="shortcut-last-week" data-range-preset="last-week">Last Week</button>
      </div>
      <div class="legend">
        <span class="legend-pill"><span class="chip-dot"></span>Mini cards show team workload by RMI estimates</span>
        <span class="legend-pill">Cards open Jira epic links in a new tab</span>
      </div>
    </section>
    <section class="gantt-wrap" id="gantt-wrap">
      <div id="gantt-root" class="gantt-root"></div>
    </section>
  </div>
  <script>
    const reportData = {payload};
    const allItems = Array.isArray(reportData.items) ? reportData.items : [];
    const teamNames = Array.isArray(reportData.team_names) ? reportData.team_names : [];
    const snapshotMeta = reportData.snapshot_meta && typeof reportData.snapshot_meta === "object" ? reportData.snapshot_meta : {{}};

    const generatedNode = document.getElementById("generated-at");
    const sourceNode = document.getElementById("source-file");
    const visibleCountNode = document.getElementById("visible-count");
    const includedStoriesNode = document.getElementById("included-stories");
    const totalStoriesNode = document.getElementById("total-stories");
    const excludedEpicNode = document.getElementById("excluded-epic");
    const excludedDatesNode = document.getElementById("excluded-dates");
    const excludedEstimateNode = document.getElementById("excluded-estimate");
    const teamSelect = document.getElementById("team-select");
    const shiftRangeBackButton = document.getElementById("shift-range-back");
    const shiftRangeForwardButton = document.getElementById("shift-range-forward");
    const fromInput = document.getElementById("from-date");
    const toInput = document.getElementById("to-date");
    const applyButton = document.getElementById("apply-range");
    const resetButton = document.getElementById("reset-range");
    const quickFilterButtons = Array.from(document.querySelectorAll("[data-range-preset]"));
    const ganttRoot = document.getElementById("gantt-root");

    const DAY_MS = 86400000;
    const DAY_PX = 16;
    const CARD_HEIGHT = 70;
    const CARD_GAP = 8;
    const TIMELINE_SIDE_PAD = 8;

    generatedNode.textContent = snapshotMeta.snapshot_utc || reportData.generated_at || "-";
    sourceNode.textContent = reportData.source_file || "-";
    includedStoriesNode.textContent = String(snapshotMeta.included_story_rows || 0);
    totalStoriesNode.textContent = String(snapshotMeta.total_story_rows || 0);
    excludedEpicNode.textContent = String(snapshotMeta.excluded_missing_epic || 0);
    excludedDatesNode.textContent = String(snapshotMeta.excluded_missing_dates || 0);
    excludedEstimateNode.textContent = String(snapshotMeta.excluded_missing_estimate || 0);

    function syncTeamOptions() {{
      const currentValue = String(teamSelect.value || "__all__");
      const optionHtml = ['<option value="__all__">All teams</option>']
        .concat(teamNames.map((teamName) => `<option value="${{safeText(teamName)}}">${{safeText(teamName)}}</option>`));
      teamSelect.innerHTML = optionHtml.join("");
      const canKeepCurrent = currentValue === "__all__" || teamNames.includes(currentValue);
      teamSelect.value = canKeepCurrent ? currentValue : "__all__";
    }}

    function parseIso(iso) {{
      if (!iso) return null;
      const d = new Date(`${{iso}}T00:00:00`);
      if (Number.isNaN(d.getTime())) return null;
      return d;
    }}

    function isoFromDate(d) {{
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${{y}}-${{m}}-${{day}}`;
    }}

    function addDays(d, days) {{
      return new Date(d.getTime() + (days * DAY_MS));
    }}

    function daysInMonth(year, monthIndex) {{
      return new Date(year, monthIndex + 1, 0).getDate();
    }}

    function addMonths(d, months) {{
      const year = d.getFullYear();
      const month = d.getMonth();
      const day = d.getDate();
      const targetMonthIndex = month + months;
      const targetYear = year + Math.floor(targetMonthIndex / 12);
      const normalizedMonth = ((targetMonthIndex % 12) + 12) % 12;
      const targetDay = Math.min(day, daysInMonth(targetYear, normalizedMonth));
      return new Date(targetYear, normalizedMonth, targetDay);
    }}

    function startOfMonth(d) {{
      return new Date(d.getFullYear(), d.getMonth(), 1);
    }}

    function endOfMonth(d) {{
      return new Date(d.getFullYear(), d.getMonth() + 1, 0);
    }}

    function startOfYear(d) {{
      return new Date(d.getFullYear(), 0, 1);
    }}

    function endOfYear(d) {{
      return new Date(d.getFullYear(), 11, 31);
    }}

    function startOfQuarter(d) {{
      const quarterStartMonth = Math.floor(d.getMonth() / 3) * 3;
      return new Date(d.getFullYear(), quarterStartMonth, 1);
    }}

    function endOfQuarter(d) {{
      const quarterStartMonth = Math.floor(d.getMonth() / 3) * 3;
      return new Date(d.getFullYear(), quarterStartMonth + 3, 0);
    }}

    function defaultRange() {{
      const today = new Date();
      const prevMonthStart = startOfMonth(new Date(today.getFullYear(), today.getMonth() - 1, 1));
      const nextMonthEnd = endOfMonth(new Date(today.getFullYear(), today.getMonth() + 1, 1));
      return {{ from: prevMonthStart, to: nextMonthEnd }};
    }}

    function overlap(aStart, aEnd, bStart, bEnd) {{
      return aStart <= bEnd && aEnd >= bStart;
    }}

    function daysBetweenInclusive(startDate, endDate) {{
      return Math.max(1, Math.floor((endDate.getTime() - startDate.getTime()) / DAY_MS) + 1);
    }}

    function pctFromDays(days, totalDays) {{
      if (totalDays <= 0) return 0;
      return (days / totalDays) * 100;
    }}

    function formatShortDate(d) {{
      return d.toLocaleDateString(undefined, {{ day: "2-digit", month: "short" }});
    }}

    function monthLabel(d) {{
      return d.toLocaleDateString(undefined, {{ month: "short", year: "numeric" }});
    }}

    function weekStartFor(dateObj) {{
      const out = new Date(dateObj.getTime());
      const day = out.getDay();
      const shift = day === 0 ? 6 : day - 1;
      out.setDate(out.getDate() - shift);
      out.setHours(0, 0, 0, 0);
      return out;
    }}

    function buildWeeks(rangeFrom, rangeTo) {{
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const weeks = [];
      let cursor = weekStartFor(rangeFrom);
      while (cursor <= rangeTo) {{
        const start = new Date(cursor.getTime());
        const end = addDays(start, 6);
        const clippedStart = start < rangeFrom ? rangeFrom : start;
        const clippedEnd = end > rangeTo ? rangeTo : end;
        const totalDays = daysBetweenInclusive(rangeFrom, rangeTo);
        const leftDays = Math.floor((clippedStart.getTime() - rangeFrom.getTime()) / DAY_MS);
        const spanDays = daysBetweenInclusive(clippedStart, clippedEnd);
        weeks.push({{
          start,
          clippedStart,
          clippedEnd,
          isCurrent: today >= start && today <= end,
          leftPct: pctFromDays(leftDays, totalDays),
          widthPct: pctFromDays(spanDays, totalDays),
          label: `Wk of ${{formatShortDate(start)}}`,
        }});
        cursor = addDays(cursor, 7);
      }}
      return weeks;
    }}

    function buildMonths(rangeFrom, rangeTo) {{
      const months = [];
      let cursor = startOfMonth(rangeFrom);
      const totalDays = daysBetweenInclusive(rangeFrom, rangeTo);
      while (cursor <= rangeTo) {{
        const monthStart = new Date(cursor.getTime());
        const monthEnd = endOfMonth(cursor);
        const clippedStart = monthStart < rangeFrom ? rangeFrom : monthStart;
        const clippedEnd = monthEnd > rangeTo ? rangeTo : monthEnd;
        const leftDays = Math.floor((clippedStart.getTime() - rangeFrom.getTime()) / DAY_MS);
        const spanDays = daysBetweenInclusive(clippedStart, clippedEnd);
        months.push({{
          label: monthLabel(monthStart),
          leftPct: pctFromDays(leftDays, totalDays),
          widthPct: pctFromDays(spanDays, totalDays),
        }});
        cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
      }}
      return months;
    }}

    function parseRangeFromInputs() {{
      const preset = defaultRange();
      const from = parseIso(fromInput.value) || preset.from;
      const to = parseIso(toInput.value) || preset.to;
      if (to < from) {{
        return {{ from: preset.from, to: preset.to }};
      }}
      return {{ from, to }};
    }}

    function injectInputs(from, to) {{
      fromInput.value = isoFromDate(from);
      toInput.value = isoFromDate(to);
    }}

    function setActiveQuickFilter(presetKey) {{
      quickFilterButtons.forEach((button) => {{
        button.classList.toggle("active", String(button.dataset.rangePreset || "") === String(presetKey || ""));
      }});
    }}

    function clearActiveQuickFilter() {{
      setActiveQuickFilter("");
    }}

    function resolvePresetRange(presetKey) {{
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      switch (presetKey) {{
        case "this-year":
          return {{ from: startOfYear(today), to: endOfYear(today) }};
        case "this-month":
          return {{ from: startOfMonth(today), to: endOfMonth(today) }};
        case "previous-month": {{
          const previousMonthDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
          return {{ from: startOfMonth(previousMonthDate), to: endOfMonth(previousMonthDate) }};
        }}
        case "this-quarter":
          return {{ from: startOfQuarter(today), to: endOfQuarter(today) }};
        case "this-week": {{
          const from = weekStartFor(today);
          return {{ from, to: addDays(from, 6) }};
        }}
        case "last-week": {{
          const currentWeekStart = weekStartFor(today);
          const from = addDays(currentWeekStart, -7);
          return {{ from, to: addDays(from, 6) }};
        }}
        default:
          return defaultRange();
      }}
    }}

    function shiftCurrentRangeByMonths(monthOffset) {{
      const currentRange = parseRangeFromInputs();
      const shiftedFrom = addMonths(currentRange.from, monthOffset);
      const shiftedTo = addMonths(currentRange.to, monthOffset);
      injectInputs(shiftedFrom, shiftedTo);
      clearActiveQuickFilter();
      render();
    }}

    function filteredItems(rangeFrom, rangeTo) {{
      const selectedTeam = String(teamSelect.value || "__all__");
      return allItems.filter((item) => {{
        const start = parseIso(item.planned_start);
        const end = parseIso(item.planned_end);
        if (!start || !end || end < start) return false;
        if (selectedTeam !== "__all__" && String(item.team_name || "") !== selectedTeam) return false;
        return overlap(start, end, rangeFrom, rangeTo);
      }});
    }}

    function stackCards(items) {{
      const sorted = [...items].sort((a, b) => {{
        if (a.planned_start !== b.planned_start) return a.planned_start < b.planned_start ? -1 : 1;
        if (a.planned_end !== b.planned_end) return a.planned_end < b.planned_end ? -1 : 1;
        return (a.epic_name || "").localeCompare(b.epic_name || "");
      }});
      const trackEnds = [];
      return sorted.map((item) => {{
        const start = parseIso(item.planned_start);
        const end = parseIso(item.planned_end);
        let track = 0;
        for (; track < trackEnds.length; track++) {{
          if (trackEnds[track] < start) break;
        }}
        if (track === trackEnds.length) {{
          trackEnds.push(end);
        }} else {{
          trackEnds[track] = end;
        }}
        return {{ ...item, _track: track }};
      }});
    }}

    function laneLoadChips(laneItems, weeks) {{
      if (!weeks.length) return "";
      const weekLoads = weeks.map((week) => {{
        let total = 0;
        for (const item of laneItems) {{
          const s = parseIso(item.planned_start);
          const e = parseIso(item.planned_end);
          if (!s || !e) continue;
          if (overlap(s, e, week.clippedStart, week.clippedEnd)) {{
            total += Number(item.planned_man_days || 0);
          }}
        }}
        return total;
      }});
      const maxLoad = Math.max(...weekLoads, 0);
      return weekLoads.map((load, i) => {{
        const pct = maxLoad > 0 ? (load / maxLoad) : 0;
        const bg = pct <= 0
          ? "#eff4f8"
          : pct < 0.35
            ? "#dbeafe"
            : pct < 0.7
              ? "#93c5fd"
              : "#2563eb";
        const pretty = Number(load.toFixed(2)).toString();
        const title = `${{weeks[i].label}}: ${{pretty}} man-days`;
        return `<span class="load-chip" style="background:${{bg}};" title="${{title}}"></span>`;
      }}).join("");
    }}

    function safeText(value) {{
      return String(value == null ? "" : value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }}

    function normalizeStatus(status) {{
      return String(status || "").trim().toLowerCase();
    }}

    function statusStyle(status) {{
      const normalized = normalizeStatus(status);
      if (!normalized) {{
        return {{ accent: "#7cb6d1", border: "rgba(15, 76, 92, 0.32)", top: "rgba(219, 238, 247, 0.82)", bottom: "rgba(233, 246, 253, 0.95)", pillBg: "#edf5fa", pillBorder: "#bfd3df", pillText: "#1f4658", label: "Unknown" }};
      }}
      if (normalized.includes("done") || normalized.includes("resolved") || normalized.includes("closed") || normalized.includes("complete")) {{
        return {{ accent: "#2f855a", border: "rgba(47, 133, 90, 0.34)", top: "rgba(220, 252, 231, 0.92)", bottom: "rgba(240, 253, 244, 0.98)", pillBg: "#dcfce7", pillBorder: "#86efac", pillText: "#166534", label: status }};
      }}
      if (normalized.includes("hold") || normalized.includes("block") || normalized.includes("stuck")) {{
        return {{ accent: "#c05621", border: "rgba(192, 86, 33, 0.34)", top: "rgba(255, 237, 213, 0.95)", bottom: "rgba(255, 247, 237, 0.98)", pillBg: "#ffedd5", pillBorder: "#fdba74", pillText: "#9a3412", label: status }};
      }}
      if (normalized.includes("progress") || normalized.includes("process") || normalized.includes("development") || normalized.includes("review") || normalized.includes("testing")) {{
        return {{ accent: "#2563eb", border: "rgba(37, 99, 235, 0.34)", top: "rgba(219, 234, 254, 0.95)", bottom: "rgba(239, 246, 255, 0.98)", pillBg: "#dbeafe", pillBorder: "#93c5fd", pillText: "#1d4ed8", label: status }};
      }}
      if (normalized.includes("todo") || normalized.includes("to do") || normalized.includes("queue") || normalized.includes("selected") || normalized.includes("backlog") || normalized.includes("open")) {{
        return {{ accent: "#6b7280", border: "rgba(107, 114, 128, 0.34)", top: "rgba(243, 244, 246, 0.96)", bottom: "rgba(249, 250, 251, 0.98)", pillBg: "#f3f4f6", pillBorder: "#d1d5db", pillText: "#374151", label: status }};
      }}
      return {{ accent: "#7c3aed", border: "rgba(124, 58, 237, 0.30)", top: "rgba(243, 232, 255, 0.95)", bottom: "rgba(250, 245, 255, 0.98)", pillBg: "#ede9fe", pillBorder: "#c4b5fd", pillText: "#6d28d9", label: status }};
    }}

    function render() {{
      const range = parseRangeFromInputs();
      injectInputs(range.from, range.to);

      const visible = filteredItems(range.from, range.to);
      visibleCountNode.textContent = String(visible.length);

      const totalDays = daysBetweenInclusive(range.from, range.to);
      const timelineWidth = Math.max(880, Math.floor(totalDays * DAY_PX));
      const months = buildMonths(range.from, range.to);
      const weeks = buildWeeks(range.from, range.to);
      const selectedTeam = String(teamSelect.value || "__all__");
      const visibleTeamNames = selectedTeam === "__all__"
        ? teamNames
        : teamNames.filter((teamName) => teamName === selectedTeam);

      if (!teamNames.length) {{
        ganttRoot.innerHTML = '<div class="empty">No team workload rows found in SQLite snapshot. Run sync_team_rmi_gantt_sqlite.py first.</div>';
        return;
      }}

      const monthBlocks = months.map((m) =>
        `<div class="month-block" style="left:${{m.leftPct}}%; width:${{m.widthPct}}%;">${{m.label}}</div>`
      ).join("");
      const weekBlocks = weeks.map((w) =>
        `<div class="week-block${{w.isCurrent ? " current-week" : ""}}" style="left:${{w.leftPct}}%; width:${{w.widthPct}}%;" title="${{w.label}}">${{w.label}}</div>`
      ).join("");
      const weekLines = weeks.map((w) => `<span class="week-line" style="left:${{w.leftPct}}%;"></span>`).join("");
      const currentWeek = weeks.find((w) => w.isCurrent);
      const currentBandHtml = currentWeek
        ? `<span class="current-week-band" style="left:${{currentWeek.leftPct}}%; width:${{currentWeek.widthPct}}%;"></span>`
        : "";

      const laneRows = visibleTeamNames.map((teamName) => {{
        const laneItems = visible.filter((item) => item.team_name === teamName);
        const stacked = stackCards(laneItems);
        const maxTrack = stacked.reduce((acc, item) => Math.max(acc, Number(item._track || 0)), -1);
        const rowCount = Math.max(1, maxTrack + 1);
        const laneHeight = Math.max(48, (rowCount * CARD_HEIGHT) + ((rowCount - 1) * CARD_GAP) + 20);
        const chips = laneLoadChips(laneItems, weeks);

        let cardsHtml = "";
        for (const item of stacked) {{
          const start = parseIso(item.planned_start);
          const end = parseIso(item.planned_end);
          if (!start || !end) continue;

          const clippedStart = start < range.from ? range.from : start;
          const clippedEnd = end > range.to ? range.to : end;
          const leftDays = Math.floor((clippedStart.getTime() - range.from.getTime()) / DAY_MS);
          const spanDays = daysBetweenInclusive(clippedStart, clippedEnd);
          const leftPx = Math.floor(leftDays * DAY_PX) + TIMELINE_SIDE_PAD;
          const widthPx = Math.max(140, Math.floor(spanDays * DAY_PX) - 4);
          const topPx = 8 + (Number(item._track || 0) * (CARD_HEIGHT + CARD_GAP));
          const title = [
            `Team: ${{item.team_name}}`,
            `Epic: ${{item.epic_key}} - ${{item.epic_name}}`,
            `Story Count: ${{item.story_count}}`,
            `Planned Hours: ${{item.planned_hours}}`,
            `Planned Man Days: ${{item.planned_man_days}}`,
            `Planned Start: ${{item.planned_start}}`,
            `Planned End: ${{item.planned_end}}`,
            `Project: ${{item.project_key || "-"}}`,
          ].join("\\n");

          const epicUrl = safeText(item.epic_url || "#");
          const epicLabel = safeText(item.epic_key || "-");
          const epicName = safeText(item.epic_name || "-");
          const storyCount = safeText(item.story_count);
          const plannedHours = safeText(item.planned_hours);
          const plannedManDays = safeText(item.planned_man_days);
          const plannedStart = safeText(item.planned_start || "-");
          const plannedEnd = safeText(item.planned_end || "-");
          const statusInfo = statusStyle(item.epic_status);
          const statusLabel = safeText(statusInfo.label);

          cardsHtml += `
            <article class="card" style="left:${{leftPx}}px; top:${{topPx}}px; width:${{widthPx}}px; --card-accent:${{statusInfo.accent}}; --card-border:${{statusInfo.border}}; --card-bg-top:${{statusInfo.top}}; --card-bg-bottom:${{statusInfo.bottom}}; --pill-bg:${{statusInfo.pillBg}}; --pill-border:${{statusInfo.pillBorder}}; --pill-text:${{statusInfo.pillText}};" title="${{safeText(title)}}">
              <a class="card-link" href="${{epicUrl}}" target="_blank" rel="noopener">
                <div class="status-pill">${{statusLabel}}</div>
                <div class="card-title">${{epicLabel}} - ${{epicName}}</div>
                <div class="card-meta">Stories: ${{storyCount}} | Hours: ${{plannedHours}} | Man-days: ${{plannedManDays}}</div>
                <div class="card-meta">Start: ${{plannedStart}} | End: ${{plannedEnd}}</div>
              </a>
            </article>
          `;
        }}

        const lineEveryDays = 7;
        const lines = [];
        for (let day = 0; day <= totalDays; day += lineEveryDays) {{
          const leftPx = Math.floor(day * DAY_PX) + TIMELINE_SIDE_PAD;
          lines.push(`<span class="grid-line" style="left:${{leftPx}}px;"></span>`);
        }}
        const linesHtml = lines.join("");

        const contentHtml = stacked.length
          ? `<div class="lane-timeline" style="width:${{timelineWidth}}px; min-height:${{laneHeight}}px;">${{linesHtml}}${{weekLines}}${{currentBandHtml}}${{cardsHtml}}</div>`
          : `<div class="lane-timeline" style="width:${{timelineWidth}}px; min-height:48px;">${{linesHtml}}${{weekLines}}${{currentBandHtml}}<div class="lane-empty">No RMIs in selected range.</div></div>`;

        return `
          <div class="grid-row">
            <div class="cell team-cell">
              <div class="team-title" title="${{safeText(teamName)}}">${{safeText(teamName)}}</div>
              <div class="team-sub">${{laneItems.length}} RMI card(s)</div>
              <div class="load-row">${{chips}}</div>
            </div>
            <div class="cell timeline-cell">${{contentHtml}}</div>
          </div>
        `;
      }}).join("");

      ganttRoot.innerHTML = laneRows
        ? `
        <div class="grid" style="width: calc(var(--team-col) + ${{timelineWidth}}px);">
          <div class="grid-row head">
            <div class="cell team-cell">Team</div>
            <div class="cell timeline-cell">
              <div class="timeline-header" style="width:${{timelineWidth}}px;">
                ${{monthBlocks}}
                ${{currentBandHtml}}
                ${{weekBlocks}}
              </div>
            </div>
          </div>
          ${{laneRows}}
        </div>
      `
        : '<div class="empty">No team workload rows found for the current team/date filters.</div>';
    }}

    syncTeamOptions();
    const defaults = defaultRange();
    injectInputs(defaults.from, defaults.to);
    clearActiveQuickFilter();
    applyButton.addEventListener("click", () => {{
      clearActiveQuickFilter();
      render();
    }});
    teamSelect.addEventListener("change", render);
    fromInput.addEventListener("change", clearActiveQuickFilter);
    toInput.addEventListener("change", clearActiveQuickFilter);
    quickFilterButtons.forEach((button) => {{
      button.addEventListener("click", () => {{
        const presetKey = String(button.dataset.rangePreset || "");
        const range = resolvePresetRange(presetKey);
        setActiveQuickFilter(presetKey);
        injectInputs(range.from, range.to);
        render();
      }});
    }});
    shiftRangeBackButton.addEventListener("click", () => {{
      shiftCurrentRangeByMonths(-1);
    }});
    shiftRangeForwardButton.addEventListener("click", () => {{
      shiftCurrentRangeByMonths(1);
    }});
    resetButton.addEventListener("click", () => {{
      teamSelect.value = "__all__";
      const d = defaultRange();
      injectInputs(d.from, d.to);
      clearActiveQuickFilter();
      render();
    }});
    render();
  </script>
<script src="shared-nav.js"></script>
</body>
</html>
"""


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    db_name = os.getenv("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", DEFAULT_CAPACITY_DB).strip() or DEFAULT_CAPACITY_DB
    output_name = os.getenv("JIRA_PHASE_GANTT_HTML_PATH", DEFAULT_OUTPUT_HTML).strip() or DEFAULT_OUTPUT_HTML

    db_path = _resolve_path(db_name, base_dir)
    output_path = _resolve_path(output_name, base_dir)
    loaded = _load_team_rmi_payload(db_path)

    payload = {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "source_file": loaded.get("source_file", str(db_path)),
        "team_names": loaded.get("team_names", []),
        "items": loaded.get("items", []),
        "snapshot_meta": loaded.get("snapshot_meta", {}),
    }
    output_path.write_text(_build_html(payload), encoding="utf-8")

    print(f"Capacity DB: {db_path}")
    print(f"Loaded team names: {len(payload['team_names'])}")
    print(f"Loaded team-RMI rows: {len(payload['items'])}")
    print(f"Report written: {output_path}")


if __name__ == "__main__":
    main()
