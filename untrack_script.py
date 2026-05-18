#!/usr/bin/env python3
import os
import subprocess
import sys

os.chdir(r"E:\JIRA SCRIPT")

# Files to untrack
files_to_untrack = [
    "__debug_page_1.png",
    "__debug_page_1_after_fix.png",
    "__debug_page_2.png",
    "__debug_page_ipp_db_fetch.png",
    "__selenium_epics_debug.png",
    "__selenium_epics_debug_after_restart.png",
    "__tmp_rnd_story.js",
    "temp_ipp.xlsx",
    "temp_ipp_source.xlsx",
    "temp_ipp.pptx",
    "inspect_page_structure.py",
    "inspect_screenshot.png",
    "_apply_rowid_patch.py",
    "compare_offline_employee_performance.py",
    "copy_dashboard.py",
    "offline_html_prepare.py",
    "canonical_report_data.py",
    "screenshot_01_initial.png",
    "screenshot_02_jan2025.png",
    "screenshot_03_jun2025.png",
    "screenshot_1_initial.png",
    "screenshot_2_after_filter.png",
    "screenshot_2_jan2026.png",
    "screenshot_3_feb2026.png",
    "screenshot_final_test.png",
    "screenshot_initial.png",
    "test_date_filter.py",
    "test_date_filter_detailed.py",
    "test_date_filter_detailed_report.py",
    "test_scorecard_changes.py",
    "test_nested_view_console.js",
    "test_results.txt",
    "startup.txt",
    "run_all.log",
    "server_restart.log",
    "server_restart_err.log",
    "ipp_phase_breakdown_2026-02-27.xlsx",
    "ipp_phase_breakdown_2026-03-02.xlsx",
    "ipp_phase_breakdown_2026-03-03.xlsx",
    "ipp_phase_breakdown_2026-03-04.xlsx",
    "ipp_phase_breakdown_2026-03-05.xlsx",
    "ipp_phase_breakdown_2026-03-10.xlsx",
    "ipp_phase_breakdown_test.xlsx",
    "jira_manually_exported_original_estimates_O2_Project.xlsx",
    "export - worklog hours on jira for all projects of march.xlsx",
    "assignee_hours_capacity.sqlite",
    "assignee_hours_capacity.db-shm",
    "assignee_hours_capacity.db-wal",
    "handover.zip",
    "Late (In ProgressOn Hold).txt",
    "cli commands.txt",
]

# Get list of tracked files
result = subprocess.run(["git", "ls-files"], capture_output=True, text=True)
tracked_files = set(result.stdout.strip().split('\n'))

print("=== CHECKING TRACKED FILES ===")
untracked_count = 0
tracked_and_untracked = []

for file in files_to_untrack:
    if file in tracked_files:
        print(f"[TRACKED] {file}")
        tracked_and_untracked.append(file)
        untracked_count += 1
    else:
        print(f"[UNTRACKED] {file}")

print(f"\n=== UNTRACKING {untracked_count} FILES ===")

# Untrack the files
for file in tracked_and_untracked:
    try:
        result = subprocess.run(["git", "rm", "--cached", file], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✓ Untracked: {file}")
        else:
            print(f"✗ Failed to untrack {file}: {result.stderr}")
    except Exception as e:
        print(f"✗ Error untracking {file}: {e}")

print("\n=== COMPLETE ===")
print(f"Untracked {untracked_count} files from git")
