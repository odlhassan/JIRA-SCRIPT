"""
Run Jira export scripts sequentially with safe handoffs between outputs/inputs.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


def _resolve_output_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return base_dir / path


def _run_step(step_name: str, script_name: str, base_dir: Path, env: dict[str, str]) -> None:
    script_path = base_dir / script_name
    if not script_path.exists():
        raise FileNotFoundError(f"Step '{step_name}' missing script: {script_path}")

    print(f"\n[{step_name}] Running {script_name}")
    command = [sys.executable, str(script_path)]
    result = subprocess.run(command, cwd=str(base_dir), env=env)
    if result.returncode != 0:
        raise RuntimeError(f"Step '{step_name}' failed with exit code {result.returncode}")
    print(f"[{step_name}] Completed")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run all Jira export scripts in sequence."
    )
    parser.add_argument(
        "--skip-work-items",
        action="store_true",
        help="Skip export_jira_work_items.py",
    )
    parser.add_argument(
        "--skip-worklogs",
        action="store_true",
        help="Skip export_jira_subtask_worklogs.py",
    )
    parser.add_argument(
        "--skip-rollup",
        action="store_true",
        help="Skip export_jira_subtask_worklog_rollup.py",
    )
    parser.add_argument(
        "--skip-nested-view",
        action="store_true",
        help="Skip export_jira_nested_view.py",
    )
    parser.add_argument(
        "--incremental",
        action="store_true",
        help="Enable smart incremental fetch (default: full fetch).",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    env = os.environ.copy()
    if args.incremental:
        env["JIRA_INCREMENTAL_DISABLE"] = "0"
    else:
        env["JIRA_INCREMENTAL_DISABLE"] = "1"
    incremental_disabled = (env.get("JIRA_INCREMENTAL_DISABLE", "1").strip() or "1") == "1"
    sync_db_path = (env.get("JIRA_SYNC_DB_PATH", "jira_sync_cache.db").strip() or "jira_sync_cache.db")

    work_items_output = env.get("JIRA_EXPORT_XLSX_PATH", "1_jira_work_items_export.xlsx").strip()
    if not work_items_output:
        work_items_output = "1_jira_work_items_export.xlsx"
    env["JIRA_EXPORT_XLSX_PATH"] = work_items_output

    # Ensure rollup reads from the same file that worklog export writes.
    worklog_output = env.get("JIRA_WORKLOG_XLSX_PATH", "2_jira_subtask_worklogs.xlsx").strip()
    if not worklog_output:
        worklog_output = "2_jira_subtask_worklogs.xlsx"
    env["JIRA_WORKLOG_XLSX_PATH"] = worklog_output
    env["JIRA_SUBTASK_WORKLOG_INPUT_XLSX_PATH"] = str(_resolve_output_path(worklog_output, base_dir))

    rollup_output = env.get("JIRA_SUBTASK_ROLLUP_XLSX_PATH", "3_jira_subtask_worklog_rollup.xlsx").strip()
    if not rollup_output:
        rollup_output = "3_jira_subtask_worklog_rollup.xlsx"
    env["JIRA_SUBTASK_ROLLUP_XLSX_PATH"] = rollup_output

    nested_view_output = env.get("JIRA_NESTED_VIEW_XLSX_PATH", "nested view.xlsx").strip()
    if not nested_view_output:
        nested_view_output = "nested view.xlsx"
    env["JIRA_NESTED_VIEW_XLSX_PATH"] = nested_view_output

    print("Starting Jira export orchestration")
    print(f"Incremental sync: {'disabled' if incremental_disabled else 'enabled'}")
    print(f"Sync DB path: {sync_db_path}")
    print(f"Work items output: {env['JIRA_EXPORT_XLSX_PATH']}")
    print(f"Worklog output: {env['JIRA_WORKLOG_XLSX_PATH']}")
    print(f"Rollup input: {env['JIRA_SUBTASK_WORKLOG_INPUT_XLSX_PATH']}")
    print(f"Rollup output: {env['JIRA_SUBTASK_ROLLUP_XLSX_PATH']}")
    print(f"Nested view output: {env['JIRA_NESTED_VIEW_XLSX_PATH']}")

    if not args.skip_worklogs:
        _run_step("subtask-worklogs", "export_jira_subtask_worklogs.py", base_dir, env)
    else:
        print("\n[subtask-worklogs] Skipped")

    if not args.skip_work_items:
        _run_step("work-items", "export_jira_work_items.py", base_dir, env)
    else:
        print("\n[work-items] Skipped")

    if not args.skip_rollup:
        if args.skip_worklogs:
            expected_input = Path(env["JIRA_SUBTASK_WORKLOG_INPUT_XLSX_PATH"])
            if not expected_input.exists():
                raise FileNotFoundError(
                    "Rollup input is missing. Run worklogs step first or set "
                    "JIRA_SUBTASK_WORKLOG_INPUT_XLSX_PATH to an existing file."
                )
        _run_step("subtask-rollup", "export_jira_subtask_worklog_rollup.py", base_dir, env)
    else:
        print("\n[subtask-rollup] Skipped")

    if not args.skip_nested_view:
        if args.skip_rollup:
            expected_rollup = _resolve_output_path(env["JIRA_SUBTASK_ROLLUP_XLSX_PATH"], base_dir)
            if not expected_rollup.exists():
                raise FileNotFoundError(
                    "Nested view input rollup is missing. Run rollup step first or set "
                    "JIRA_SUBTASK_ROLLUP_XLSX_PATH to an existing file."
                )
        _run_step("nested-view", "export_jira_nested_view.py", base_dir, env)
    else:
        print("\n[nested-view] Skipped")

    print("\nAll requested export steps completed successfully.")


if __name__ == "__main__":
    main()
