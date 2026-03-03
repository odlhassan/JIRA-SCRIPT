from __future__ import annotations

import os
import shutil
from pathlib import Path


def _resolve_output_path(base_dir: Path) -> Path:
    raw_value = (os.getenv("JIRA_PLANNED_ACTUAL_TABLE_VIEW_HTML_PATH", "planned_actual_table_view.html") or "").strip()
    path = Path(raw_value or "planned_actual_table_view.html")
    if not path.is_absolute():
        path = base_dir / path
    return path


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    source = base_dir / "report_html" / "planned_actual_table_view.html"
    output = _resolve_output_path(base_dir)

    if not source.exists() or not source.is_file():
        raise FileNotFoundError(f"Source template missing: {source}")

    output.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == output.resolve():
        print(f"[planned-actual-table-view-html] Up-to-date at {output}")
        return

    shutil.copy2(source, output)
    print(f"[planned-actual-table-view-html] Wrote {output}")


if __name__ == "__main__":
    main()
