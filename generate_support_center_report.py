from __future__ import annotations

import os
import shutil
from pathlib import Path


def _resolve_output_path(base_dir: Path) -> Path:
    raw_value = (os.getenv("JIRA_SUPPORT_CENTER_HTML_PATH", "support_center_report.html") or "").strip()
    path = Path(raw_value or "support_center_report.html")
    if not path.is_absolute():
        path = base_dir / path
    return path


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    source = base_dir / "support_center_report.html"
    published = base_dir / "report_html" / "support_center_report.html"

    if not source.exists() or not source.is_file():
        raise FileNotFoundError(f"Source shell missing: {source}")

    output = _resolve_output_path(base_dir)
    output.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != output.resolve():
        shutil.copy2(source, output)
        print(f"[support-center-html] Wrote {output}")
    else:
        print(f"[support-center-html] Canonical up-to-date at {output}")

    published.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != published.resolve():
        shutil.copy2(source, published)
        print(f"[support-center-html] Published {published}")


if __name__ == "__main__":
    main()
