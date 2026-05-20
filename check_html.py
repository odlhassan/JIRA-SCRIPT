from pathlib import Path


HTML_PATH = Path("report_html") / "team_capacity_planner.html"

REQUIRED_MARKERS = {
    "Modal HTML present": 'id="jira-sync-modal"',
    "Checkboxes present": "jira-row-chk",
    "Toolbar present": "jira-sel-toolbar",
    "Select Unsynced btn": "sel-unsynced-btn",
    "_jiraSyncSelectedIds": "_jiraSyncSelectedIds",
    "ids sent in push": "JSON.stringify({ ids })",
    "Stats toggle present": 'id="stat-unit-toggle"',
    "Hours option present": 'data-unit="hours"',
    "Days option present": 'data-unit="days"',
    "Subtask-only planned marker": "Assigned subtask estimates only",
    "Subtask planned payload field": "subtask_planned_hours",
}


def main() -> int:
    content = HTML_PATH.read_text(encoding="utf-8")
    print(f"File size: {len(content)} chars")
    missing = []
    for label, marker in REQUIRED_MARKERS.items():
        present = marker in content
        print(f"{label}: {present}")
        if not present:
            missing.append(label)
    if missing:
        print("Missing required Team Capacity Planner HTML markers: " + ", ".join(missing))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
