from __future__ import annotations

import os
import shutil
from pathlib import Path

LEGACY_OUTPUT = "planned_vs_dispensed_report.html"
CANONICAL_OUTPUT = "approved_vs_planned_hours_report.html"

def _resolve_output_path(base_dir: Path) -> Path:
    raw_value = (os.getenv("JIRA_PLANNED_VS_DISPENSED_HTML_PATH", LEGACY_OUTPUT) or "").strip()
    path = Path(raw_value or LEGACY_OUTPUT)
    if not path.is_absolute():
        path = base_dir / path
    return path


def _canonical_alias_path(output: Path) -> Path:
    return output.with_name(CANONICAL_OUTPUT)


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    source = base_dir / "report_html" / CANONICAL_OUTPUT
    if not source.exists() or not source.is_file():
        source = base_dir / "report_html" / LEGACY_OUTPUT
    output = _resolve_output_path(base_dir)
    canonical_output = _canonical_alias_path(output)

    if not source.exists() or not source.is_file():
        raise FileNotFoundError(f"Source template missing: {source}")

    output.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == output.resolve():
        print(f"[planned-vs-dispensed-html] Up-to-date at {output}")
    else:
        shutil.copy2(source, output)
        print(f"[planned-vs-dispensed-html] Wrote {output}")

    if canonical_output.resolve() != output.resolve():
        shutil.copy2(output, canonical_output)
        print(f"[planned-vs-dispensed-html] Wrote {canonical_output}")


if __name__ == "__main__":
    main()
