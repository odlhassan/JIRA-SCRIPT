#!/usr/bin/env python
import sqlite3
import json
from pathlib import Path

print("=" * 60)
print("DEBUG 1: Checking jira_sync_cache.db")
print("=" * 60)

db = Path("E:/JIRA SCRIPT/jira_sync_cache.db")
with sqlite3.connect(db) as c:
    # Check if table exists
    tables = c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
    print("Tables:", [t[0] for t in tables])
    
    # Check canonical_assignee_period_hours if it exists
    try:
        count = c.execute("SELECT COUNT(*) FROM canonical_assignee_period_hours").fetchone()
        print("\ncanonical_assignee_period_hours rows:", count)
        sample = c.execute("SELECT * FROM canonical_assignee_period_hours LIMIT 5").fetchall()
        print("Sample:", sample)
        # Check month codes available
        months = c.execute("SELECT DISTINCT period_value FROM canonical_assignee_period_hours WHERE period_type='month' ORDER BY period_value DESC LIMIT 10").fetchall()
        print("Available months:", months)
    except Exception as e:
        print("Error:", e)

print("\n" + "=" * 60)
print("DEBUG 2: Checking assignee_hours_capacity.db")
print("=" * 60)

db2 = Path("E:/JIRA SCRIPT/assignee_hours_capacity.db")
with sqlite3.connect(db2) as c:
    try:
        row = c.execute("SELECT last_success_run_id FROM canonical_refresh_state WHERE id=1").fetchone()
        print("last_success_run_id:", row)
    except Exception as e:
        print("Error:", e)

print("\nDebug complete!")
