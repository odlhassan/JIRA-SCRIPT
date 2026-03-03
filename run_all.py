"""
Run the full pipeline in one command:
1) Generate export Excel files
2) Sync team-RMI gantt snapshot tables into SQLite
3) Build nested-view HTML report
4) Build missed-entries HTML report
5) Build assignee-hours HTML + summary Excel report
6) Build dedicated RLT leave report artifacts
7) Build planned leaves calendar HTML report
8) Build RnD data story HTML report (includes project-wise epic planned-hours page)
9) Build Planned RMIs report HTML
10) Build main gantt-chart HTML report
11) Build phase-owner RMI gantt-chart HTML report
12) Build Employee Performance report HTML
13) Rebuild dashboard.html
14) Build IPP phase breakdown export
15) Build IPP Meeting dashboard HTML
16) Move generated report HTML files into the designated folder
17) Start local report server (unless --no-server is used)
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import traceback
import uuid
from datetime import datetime, timezone
from pathlib import Path

from report_server import run_report_server, sync_report_html

RUN_STATE_FILE = "run_all_state.json"
RUN_LOG_FILE = "run_all.log"


def _run_step(step_name: str, script_name: str, base_dir: Path, extra_args: list[str] | None = None, env: dict[str, str] | None = None) -> None:
    script_path = base_dir / script_name
    if not script_path.exists():
        raise FileNotFoundError(f"Missing script for '{step_name}': {script_path}")

    print(f"\n[{step_name}] Running {script_name}")
    command = [sys.executable, str(script_path)]
    if extra_args:
        command.extend(extra_args)
    result = subprocess.run(command, cwd=str(base_dir), env=env)
    if result.returncode != 0:
        raise RuntimeError(f"Step '{step_name}' failed with exit code {result.returncode}")
    print(f"[{step_name}] Completed")


def _move_report_html(base_dir: Path, folder_raw: str) -> None:
    moved = sync_report_html(base_dir, folder_raw)
    if moved == 0:
        print("[report-html-sync] No report HTML files were found to move.")
    else:
        print(f"[report-html-sync] Total moved: {moved}")


def _ensure_unified_nav_assets(base_dir: Path) -> None:
    """
    Keep shared unified-nav assets available in base_dir so generators that
    output HTML there can always reference them.
    """
    for asset in ("shared-nav.css", "shared-nav.js"):
        root_asset = base_dir / asset
        report_html_asset = base_dir / "report_html" / asset
        if root_asset.exists():
            continue
        if report_html_asset.exists():
            shutil.copy2(str(report_html_asset), str(root_asset))
            print(f"[unified-nav] Restored missing root asset: {root_asset.name}")


def _serve_report_html(base_dir: Path, folder_raw: str, host: str, port: int) -> None:
    try:
        run_report_server(base_dir=base_dir, folder_raw=folder_raw, host=host, port=port)
    except KeyboardInterrupt:
        print("\n[server] Stopped by user.")


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _state_path(base_dir: Path) -> Path:
    return base_dir / RUN_STATE_FILE


def _log_path(base_dir: Path) -> Path:
    return base_dir / RUN_LOG_FILE


def _append_log(base_dir: Path, level: str, message: str, run_id: str = "", step: str = "") -> None:
    line = {
        "timestamp_utc": _utc_now_iso(),
        "level": level,
        "run_id": run_id,
        "step": step,
        "message": message,
    }
    with _log_path(base_dir).open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(line, ensure_ascii=True) + "\n")


def _load_state(base_dir: Path) -> dict | None:
    path = _state_path(base_dir)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_state(base_dir: Path, state: dict) -> None:
    _state_path(base_dir).write_text(json.dumps(state, indent=2, ensure_ascii=True), encoding="utf-8")


def _prompt_resume_choice(previous_state: dict) -> str:
    prev_run_id = str(previous_state.get("run_id", "")).strip() or "unknown"
    prev_status = str(previous_state.get("status", "")).strip() or "unknown"
    prev_step = str(previous_state.get("current_step", "")).strip() or "n/a"
    completed = previous_state.get("completed_steps") or []
    completed_count = len([step for step in completed if str(step).strip()])
    print("\nDetected an incomplete previous pipeline run:")
    print(f"- Run ID: {prev_run_id}")
    print(f"- Last status: {prev_status}")
    print(f"- Last step in progress: {prev_step}")
    print(f"- Steps completed successfully: {completed_count}")
    while True:
        choice = input("Choose action: [C]ontinue from last successful step or [F]resh run? ").strip().lower()
        if choice in {"c", "continue"}:
            return "continue"
        if choice in {"f", "fresh"}:
            return "fresh"
        print("Please enter C or F.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the full reporting/export pipeline.")
    parser.add_argument(
        "--skip-phase-rmi-gantt",
        action="store_true",
        help="Skip generate_phase_rmi_gantt_html.py",
    )
    parser.add_argument(
        "--skip-ipp-phase-export",
        action="store_true",
        help="Skip export_ipp_phase_breakdown.py",
    )
    parser.add_argument(
        "--skip-ipp-dashboard",
        action="store_true",
        help="Skip generate_ipp_meeting_dashboard.py",
    )
    parser.add_argument(
        "--report-html-dir",
        default=os.getenv("REPORT_HTML_DIR", "report_html"),
        help="Designated folder where generated report HTML files are moved.",
    )
    parser.add_argument(
        "--no-server",
        action="store_true",
        help="Do not start local report server after pipeline completes.",
    )
    parser.add_argument(
        "--host",
        default=os.getenv("REPORT_HOST", "127.0.0.1"),
        help="Host for local report server.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("REPORT_PORT", os.getenv("PORT", "8000"))),
        help="Port for local report server.",
    )
    parser.add_argument(
        "--resume-policy",
        choices=["ask", "continue", "fresh"],
        default="ask",
        help="How to handle an incomplete prior run: ask, continue, or fresh.",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    previous_state = _load_state(base_dir)
    if previous_state and previous_state.get("status") == "running":
        previous_state = dict(previous_state)
        previous_state["status"] = "interrupted"
        previous_state["ended_at_utc"] = _utc_now_iso()
        previous_state["error"] = "Previous run ended unexpectedly (process exited before completion)."
        _save_state(base_dir, previous_state)
        _append_log(
            base_dir,
            "ERROR",
            "Previous run marked interrupted on startup due to missing completion marker.",
            run_id=str(previous_state.get("run_id", "")),
            step=str(previous_state.get("current_step", "")),
        )

    full_sync_env = os.environ.copy()
    full_sync_env["JIRA_INCREMENTAL_DISABLE"] = "1"

    pipeline_steps: list[dict] = [
        {"name": "exports", "type": "script", "script": "run_all_exports.py", "extra_args": ["--skip-work-items"], "env": None},
        {"name": "work-items-full-sync", "type": "script", "script": "export_jira_work_items.py", "extra_args": None, "env": full_sync_env},
        {"name": "team-rmi-gantt-sqlite-sync", "type": "script", "script": "sync_team_rmi_gantt_sqlite.py", "extra_args": None, "env": None},
        {"name": "nested-view-html", "type": "script", "script": "generate_nested_view_html.py", "extra_args": None, "env": None},
        {"name": "missed-entries-html", "type": "script", "script": "generate_missed_entries_html.py", "extra_args": None, "env": None},
        {"name": "assignee-hours-html", "type": "script", "script": "generate_assignee_hours_report.py", "extra_args": None, "env": None},
        {"name": "rlt-leave-report", "type": "script", "script": "generate_rlt_leave_report.py", "extra_args": None, "env": None},
        {"name": "leaves-planned-calendar-html", "type": "script", "script": "generate_leaves_planned_calendar_html.py", "extra_args": None, "env": None},
        {"name": "rnd-data-story-html", "type": "script", "script": "generate_rnd_data_story.py", "extra_args": None, "env": None},
        {"name": "planned-rmis-html", "type": "script", "script": "generate_planned_rmis_html.py", "extra_args": None, "env": None},
        {"name": "gantt-chart-html", "type": "script", "script": "generate_gantt_chart_html.py", "extra_args": None, "env": None},
        {"name": "planned-actual-table-view-html", "type": "script", "script": "generate_planned_actual_table_view.py", "extra_args": None, "env": None},
    ]
    if not args.skip_phase_rmi_gantt:
        pipeline_steps.append({"name": "phase-rmi-gantt-html", "type": "script", "script": "generate_phase_rmi_gantt_html.py", "extra_args": None, "env": None})
    else:
        print("\n[phase-rmi-gantt-html] Skipped")
    pipeline_steps.append({"name": "employee-performance-html", "type": "script", "script": "generate_employee_performance_report.py", "extra_args": None, "env": None})
    pipeline_steps.append({"name": "dashboard", "type": "script", "script": "fetch_jira_dashboard.py", "extra_args": None, "env": None})
    if not args.skip_ipp_phase_export:
        pipeline_steps.append({"name": "ipp-phase-export", "type": "script", "script": "export_ipp_phase_breakdown.py", "extra_args": None, "env": None})
    else:
        print("\n[ipp-phase-export] Skipped")
    if not args.skip_ipp_dashboard:
        pipeline_steps.append({"name": "ipp-meeting-dashboard", "type": "script", "script": "generate_ipp_meeting_dashboard.py", "extra_args": None, "env": None})
    else:
        print("\n[ipp-meeting-dashboard] Skipped")
    pipeline_steps.append({"name": "unified-nav-assets", "type": "func", "func": "_ensure_unified_nav_assets"})
    pipeline_steps.append({"name": "report-html-sync", "type": "func", "func": "_move_report_html"})

    resume_choice = "fresh"
    carried_completed: list[str] = []
    if previous_state and previous_state.get("status") in {"failed", "interrupted"}:
        if args.resume_policy == "ask":
            if sys.stdin.isatty():
                resume_choice = _prompt_resume_choice(previous_state)
            else:
                print("Incomplete previous run detected but terminal is non-interactive; starting fresh.")
                resume_choice = "fresh"
        else:
            resume_choice = args.resume_policy
        _append_log(
            base_dir,
            "INFO",
            f"Incomplete previous run detected. resume_policy={args.resume_policy}, selected={resume_choice}",
            run_id=str(previous_state.get("run_id", "")),
            step=str(previous_state.get("current_step", "")),
        )
        if resume_choice == "continue":
            prev_completed = [str(step).strip() for step in (previous_state.get("completed_steps") or []) if str(step).strip()]
            current_step_names = [str(step.get("name", "")).strip() for step in pipeline_steps]
            idx = 0
            for step_name in prev_completed:
                if idx < len(current_step_names) and step_name == current_step_names[idx]:
                    idx += 1
                else:
                    break
            carried_completed = current_step_names[:idx]
            print(f"Resuming from step {idx + 1} of {len(pipeline_steps)}.")

    run_id = uuid.uuid4().hex
    run_state = {
        "run_id": run_id,
        "status": "running",
        "started_at_utc": _utc_now_iso(),
        "ended_at_utc": None,
        "current_step": None,
        "completed_steps": carried_completed.copy(),
        "error": None,
        "resume_choice": resume_choice,
        "resumed_from_run_id": str(previous_state.get("run_id", "")) if previous_state else "",
        "args": vars(args),
    }
    _save_state(base_dir, run_state)
    _append_log(base_dir, "INFO", "Pipeline run started.", run_id=run_id)

    print("Starting full update pipeline")
    start_index = len(run_state["completed_steps"])
    try:
        for step in pipeline_steps[start_index:]:
            step_name = str(step["name"])
            run_state["current_step"] = step_name
            _save_state(base_dir, run_state)
            _append_log(base_dir, "INFO", "Step started.", run_id=run_id, step=step_name)

            if step["type"] == "script":
                _run_step(
                    step_name,
                    str(step["script"]),
                    base_dir,
                    extra_args=step.get("extra_args"),
                    env=step.get("env"),
                )
            elif step["type"] == "func":
                if step["func"] == "_ensure_unified_nav_assets":
                    _ensure_unified_nav_assets(base_dir)
                elif step["func"] == "_move_report_html":
                    _move_report_html(base_dir, args.report_html_dir)
                else:
                    raise RuntimeError(f"Unknown internal step handler: {step['func']}")
            else:
                raise RuntimeError(f"Unsupported step type: {step['type']}")

            run_state["completed_steps"].append(step_name)
            run_state["error"] = None
            _save_state(base_dir, run_state)
            _append_log(base_dir, "INFO", "Step completed.", run_id=run_id, step=step_name)
    except KeyboardInterrupt:
        run_state["status"] = "interrupted"
        run_state["ended_at_utc"] = _utc_now_iso()
        run_state["error"] = "Interrupted by user."
        _save_state(base_dir, run_state)
        _append_log(base_dir, "ERROR", "Pipeline interrupted by user.", run_id=run_id, step=str(run_state.get("current_step", "")))
        raise
    except Exception as exc:
        run_state["status"] = "failed"
        run_state["ended_at_utc"] = _utc_now_iso()
        run_state["error"] = f"{type(exc).__name__}: {exc}"
        run_state["traceback"] = traceback.format_exc()
        _save_state(base_dir, run_state)
        _append_log(
            base_dir,
            "ERROR",
            f"Pipeline failed unexpectedly: {type(exc).__name__}: {exc}",
            run_id=run_id,
            step=str(run_state.get("current_step", "")),
        )
        raise

    run_state["status"] = "completed"
    run_state["ended_at_utc"] = _utc_now_iso()
    run_state["current_step"] = None
    run_state["error"] = None
    _save_state(base_dir, run_state)
    _append_log(base_dir, "INFO", "Pipeline run completed successfully.", run_id=run_id)
    print("\nFull pipeline completed successfully.")
    if args.no_server:
        print("[server] Skipped (--no-server).")
        return
    _serve_report_html(base_dir, args.report_html_dir, args.host, args.port)


if __name__ == "__main__":
    main()
