import sqlite3, json, calendar
from pathlib import Path
from datetime import date

BASE = Path("E:/JIRA SCRIPT")
sync_db  = BASE / "jira_sync_cache.db"
settings_db = BASE / "assignee_hours_capacity.db"

today = date.today()
month_start = today.strftime("%Y-%m-01")
last_day = calendar.monthrange(today.year, today.month)[1]
month_end = today.strftime(f"%Y-%m-{last_day:02d}")
print(f"Current month window: {month_start} → {month_end}\n")

# ── 1. run_id from settings DB ──────────────────────────────────────────────
print("=== 1. run_id (assignee_hours_capacity.db) ===")
run_id = None
with sqlite3.connect(settings_db) as c:
    try:
        row = c.execute("SELECT last_success_run_id FROM canonical_refresh_state WHERE id=1").fetchone()
        run_id = row[0] if row else None
        print("run_id:", run_id)
    except Exception as e:
        print("ERROR:", e)

if not run_id:
    print("No run_id — aborting further checks")
    exit(1)

# ── 2. Does that run_id exist in canonical_issues? ───────────────────────────
print("\n=== 2. canonical_issues run_id match ===")
with sqlite3.connect(sync_db) as c:
    total = c.execute("SELECT COUNT(*) FROM canonical_issues").fetchone()[0]
    matched = c.execute("SELECT COUNT(*) FROM canonical_issues WHERE run_id=?", (run_id,)).fetchone()[0]
    all_runs = [r[0] for r in c.execute("SELECT DISTINCT run_id FROM canonical_issues ORDER BY run_id DESC LIMIT 5").fetchall()]
    print(f"Total canonical_issues rows: {total}")
    print(f"Rows matching run_id '{run_id}': {matched}")
    print(f"Top 5 run_ids in canonical_issues: {all_runs}")

# ── 3. issue_type values for that run_id ────────────────────────────────────
print("\n=== 3. issue_type distribution (run_id match) ===")
with sqlite3.connect(sync_db) as c:
    rows = c.execute(
        "SELECT issue_type, COUNT(*) FROM canonical_issues WHERE run_id=? GROUP BY issue_type ORDER BY 2 DESC",
        (run_id,)
    ).fetchall()
    for itype, cnt in rows:
        print(f"  '{itype}': {cnt}")

# ── 4. Subtask rows overlapping current month ───────────────────────────────
print("\n=== 4. Subtask rows overlapping current month ===")
with sqlite3.connect(sync_db) as c:
    rows = c.execute("""
        SELECT DISTINCT assignee,
               UPPER(COALESCE(NULLIF(TRIM(project_key),''), SUBSTR(issue_key,1,INSTR(issue_key,'-')-1))) AS pk,
               start_date, due_date, status
        FROM canonical_issues
        WHERE run_id=?
          AND (LOWER(issue_type) LIKE '%sub-task%' OR LOWER(issue_type) LIKE '%subtask%')
          AND assignee != ''
          AND LOWER(status) NOT IN ('done','resolved','closed','cancelled','rejected')
        LIMIT 20
    """, (run_id,)).fetchall()
    print(f"Active subtasks (any date, first 20): {len(rows)}")
    for r in rows[:10]:
        print(f"  {r}")

    # With date filter
    rows_dated = c.execute("""
        SELECT DISTINCT assignee,
               UPPER(COALESCE(NULLIF(TRIM(project_key),''), SUBSTR(issue_key,1,INSTR(issue_key,'-')-1))) AS pk
        FROM canonical_issues
        WHERE run_id=?
          AND (LOWER(issue_type) LIKE '%sub-task%' OR LOWER(issue_type) LIKE '%subtask%')
          AND assignee != ''
          AND LOWER(status) NOT IN ('done','resolved','closed','cancelled','rejected')
          AND (
            (start_date != '' AND start_date <= ? AND (due_date = '' OR due_date >= ?))
            OR (start_date != '' AND start_date >= ? AND start_date <= ?)
            OR (due_date  != '' AND due_date  >= ? AND due_date  <= ?)
            OR (start_date = '' AND due_date = '')
          )
        ORDER BY assignee, pk
    """, (run_id, month_end, month_start, month_start, month_end, month_start, month_end)).fetchall()
    print(f"\nSubtasks matching date filter (current month + undated): {len(rows_dated)}")
    for r in rows_dated[:10]:
        print(f"  {r}")

# ── 5. Sample date ranges of subtasks ───────────────────────────────────────
print("\n=== 5. Sample subtask date ranges ===")
with sqlite3.connect(sync_db) as c:
    rows = c.execute("""
        SELECT start_date, due_date, COUNT(*) as cnt
        FROM canonical_issues
        WHERE run_id=?
          AND (LOWER(issue_type) LIKE '%sub-task%' OR LOWER(issue_type) LIKE '%subtask%')
          AND assignee != ''
        GROUP BY start_date, due_date
        ORDER BY start_date DESC LIMIT 15
    """, (run_id,)).fetchall()
    for r in rows:
        print(f"  start={r[0]!r:20s}  due={r[1]!r:20s}  count={r[2]}")

# ── 6. API response ──────────────────────────────────────────────────────────
print("\n=== 6. API /api/seating/data ===")
try:
    import urllib.request
    with urllib.request.urlopen("http://127.0.0.1:3000/api/seating/data", timeout=5) as r:
        data = json.loads(r.read())
        print(f"_pa_count: {data.get('_pa_count')}")
        print(f"_pa_error: {data.get('_pa_error')}")
        pa = data.get("project_assignments", {})
        print(f"project_assignments entries: {len(pa)}")
        for k, v in list(pa.items())[:5]:
            print(f"  {k!r} -> {v}")
except Exception as e:
    print("API call failed:", e)
