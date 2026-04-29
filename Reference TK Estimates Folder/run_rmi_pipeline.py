from __future__ import annotations

import argparse
from pathlib import Path

from extract_rmi_jira_to_sqlite import (
    DEFAULT_DB_PATH,
    DEFAULT_SHEET_FILTER,
    DEFAULT_WORKBOOK_PATH,
    load_env_config,
    run_extraction,
)
from generate_rmi_gantt_html import DEFAULT_HTML_PATH, generate_html_report


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the full RMI pipeline: extract Jira data to SQLite, then generate the HTML gantt report."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH, help="Source Excel workbook path")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH, help="SQLite output path")
    parser.add_argument("--html", type=Path, default=DEFAULT_HTML_PATH, help="HTML output path")
    parser.add_argument(
        "--sheet-contains",
        default=DEFAULT_SHEET_FILTER,
        help="Process only sheets whose names contain this text",
    )
    parser.add_argument("--limit", type=int, default=None, help="Optional row limit for safe dry-runs")
    return parser


def run_pipeline(
    workbook_path: Path,
    db_path: Path,
    html_path: Path,
    sheet_contains: str = DEFAULT_SHEET_FILTER,
    limit: int | None = None,
) -> dict[str, object]:
    def report(message: str) -> None:
        print(f"[RMI Pipeline] {message}", flush=True)

    report("Loading Jira environment configuration")
    env_config = load_env_config()
    report(
        f"Starting extraction with workbook={workbook_path}, db={db_path}, "
        f"sheet_contains={sheet_contains!r}, limit={limit if limit is not None else 'None'}"
    )
    extraction_summary = run_extraction(
        workbook_path=workbook_path,
        db_path=db_path,
        env_config=env_config,
        sheet_contains=sheet_contains,
        limit=limit,
        progress_callback=report,
    )
    report(f"Starting HTML generation from {db_path} to {html_path}")
    generated_html = generate_html_report(db_path, html_path)
    report(f"HTML generation complete: {generated_html}")
    return {
        "env_path": env_config["ENV_PATH"],
        "db_path": db_path,
        "html_path": generated_html,
        "summary": extraction_summary,
    }


def main() -> None:
    args = build_arg_parser().parse_args()
    result = run_pipeline(
        workbook_path=args.workbook.resolve(),
        db_path=args.db.resolve(),
        html_path=args.html.resolve(),
        sheet_contains=args.sheet_contains,
        limit=args.limit,
    )
    summary = result["summary"]
    print(f"Using Jira config from: {result['env_path']}")
    print(f"Workbook: {args.workbook.resolve()}")
    print(f"SQLite DB: {result['db_path']}")
    print(f"HTML: {result['html_path']}")
    print(
        "Summary: "
        f"eligible_rows={summary['eligible_rows']}, "
        f"epics_fetched={summary['epics_fetched']}, "
        f"stories_fetched={summary['stories_fetched']}, "
        f"descendants_fetched={summary['descendants_fetched']}, "
        f"worklogs_fetched={summary['worklogs_fetched']}, "
        f"errors={summary['errors']}"
    )


if __name__ == "__main__":
    main()
