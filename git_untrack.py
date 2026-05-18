import subprocess, os

cwd = r"E:\JIRA SCRIPT"

def git(args, check=False):
    result = subprocess.run(
        ["git"] + args,
        cwd=cwd, capture_output=True, text=True
    )
    print(f"git {' '.join(args[:3])}: returncode={result.returncode}")
    if result.stdout.strip():
        print("  OUT:", result.stdout.strip()[:300])
    if result.stderr.strip():
        print("  ERR:", result.stderr.strip()[:300])
    return result

# Step 1: Get currently tracked files
tracked_result = git(["ls-files"])
tracked = set(tracked_result.stdout.strip().split("\n"))
print(f"\nTotal tracked files: {len(tracked)}\n")

# Step 2: Files to untrack
files_to_untrack = [
    "__debug_page_1.png", "__debug_page_1_after_fix.png", "__debug_page_2.png",
    "__debug_page_ipp_db_fetch.png", "__selenium_epics_debug.png",
    "__selenium_epics_debug_after_restart.png", "__tmp_rnd_story.js",
    "temp_ipp.xlsx", "temp_ipp_source.xlsx", "temp_ipp.pptx",
    "inspect_page_structure.py", "inspect_screenshot.png",
    "_apply_rowid_patch.py", "compare_offline_employee_performance.py",
    "copy_dashboard.py", "offline_html_prepare.py", "canonical_report_data.py",
    "screenshot_01_initial.png", "screenshot_02_jan2025.png", "screenshot_03_jun2025.png",
    "screenshot_1_initial.png", "screenshot_2_after_filter.png", "screenshot_2_jan2026.png",
    "screenshot_3_feb2026.png", "screenshot_final_test.png", "screenshot_initial.png",
    "test_date_filter.py", "test_date_filter_detailed.py", "test_date_filter_detailed_report.py",
    "test_scorecard_changes.py", "test_nested_view_console.js", "test_results.txt",
    "startup.txt", "run_all.log", "server_restart.log", "server_restart_err.log",
    "ipp_phase_breakdown_2026-02-27.xlsx", "ipp_phase_breakdown_2026-03-02.xlsx",
    "ipp_phase_breakdown_2026-03-03.xlsx", "ipp_phase_breakdown_2026-03-04.xlsx",
    "ipp_phase_breakdown_2026-03-05.xlsx", "ipp_phase_breakdown_2026-03-10.xlsx",
    "ipp_phase_breakdown_test.xlsx",
    "jira_manually_exported_original_estimates_O2_Project.xlsx",
    "export - worklog hours on jira for all projects of march.xlsx",
    "assignee_hours_capacity.sqlite", "assignee_hours_capacity.db-shm",
    "assignee_hours_capacity.db-wal", "handover.zip",
    "Late (In ProgressOn Hold).txt", "cli commands.txt",
]

# Step 3: Untrack each file that is currently tracked
untracked = []
skipped = []
for f in files_to_untrack:
    if f in tracked:
        result = git(["rm", "--cached", f])
        if result.returncode == 0:
            untracked.append(f)
        else:
            print(f"  FAILED to untrack: {f}")
    else:
        skipped.append(f)
        print(f"  SKIP (not tracked): {f}")

print(f"\nUntracked {len(untracked)} files, skipped {len(skipped)} (not tracked)")

# Step 4: Stage .gitignore
git(["add", ".gitignore"])

# Step 5: Commit
if untracked:
    commit_msg = "chore: untrack unused debug/temp/artifact files from git\n\nCo-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"
    git(["commit", "-m", commit_msg])
    print("\nCommit complete.")
else:
    print("\nNo files were tracked, nothing to commit (gitignore already up to date).")

# Step 6: Show final status
git(["status", "--short"])
