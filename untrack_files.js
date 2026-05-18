const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const cwd = 'E:\\JIRA SCRIPT';

// Files to untrack
const filesToUntrack = [
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
];

try {
    process.chdir(cwd);
    console.log("=== CHECKING TRACKED FILES ===");
    
    const lsFilesOutput = execSync('git ls-files', { encoding: 'utf8' });
    const trackedFiles = new Set(lsFilesOutput.trim().split('\n'));
    
    let untrackCount = 0;
    const trackedToRemove = [];
    
    for (const file of filesToUntrack) {
        if (trackedFiles.has(file)) {
            console.log(`[TRACKED] ${file}`);
            trackedToRemove.push(file);
            untrackCount++;
        } else {
            console.log(`[UNTRACKED] ${file}`);
        }
    }
    
    console.log(`\n=== UNTRACKING ${untrackCount} FILES ===`);
    
    for (const file of trackedToRemove) {
        try {
            execSync(`git rm --cached "${file}"`, { encoding: 'utf8' });
            console.log(`✓ Untracked: ${file}`);
        } catch (e) {
            console.log(`✗ Failed to untrack ${file}: ${e.message}`);
        }
    }
    
    console.log("\n=== STAGING .gitignore ===");
    execSync('git add .gitignore', { encoding: 'utf8' });
    console.log("✓ .gitignore staged");
    
    console.log("\n=== COMMITTING CHANGES ===");
    const commitMessage = "chore: untrack unused debug/temp/artifact files from git\n\nCo-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>";
    execSync(`git commit -m "${commitMessage.replace(/"/g, '\\"')}"`, { encoding: 'utf8' });
    console.log("✓ Changes committed");
    
    console.log(`\n=== COMPLETE ===`);
    console.log(`Successfully untracked ${untrackCount} files and updated .gitignore`);
    
} catch (error) {
    console.error("Error:", error.message);
    process.exit(1);
}
