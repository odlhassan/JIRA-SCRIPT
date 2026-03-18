"""Tests for Prepare Offline HTML (offline_html_prepare module)."""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from offline_html_prepare import (
    DEFAULT_OFFLINE_REPORTS,
    create_zip_of_folder,
    get_default_report_keys,
    get_default_reports_for_ui,
    run_prepare_job,
    _extract_report_data_from_html,
    _extract_json_object_after,
    _filter_rows_by_date,
    _inject_embedded_data_into_html,
    _make_fetch_override_script,
)


class TestOfflinePrepareConfig(unittest.TestCase):
    def test_default_report_keys(self):
        keys = get_default_report_keys()
        self.assertIsInstance(keys, list)
        self.assertGreaterEqual(len(keys), 10)
        self.assertIn("nested_view_report", keys)
        self.assertIn("missed_entries", keys)
        self.assertIn("planned_vs_dispensed_report", keys)

    def test_default_reports_for_ui(self):
        reports = get_default_reports_for_ui()
        self.assertIsInstance(reports, list)
        self.assertEqual(len(reports), len(DEFAULT_OFFLINE_REPORTS))
        for r in reports:
            self.assertIn("key", r)
            self.assertIn("file", r)
            self.assertIn("title", r)


class TestExtractReportData(unittest.TestCase):
    def test_extract_report_data_single_line(self):
        html = '<html><script>const reportData = {"rows": [{"id": 1}], "generated_at": "2026-01-01"};</script></html>'
        out = _extract_report_data_from_html(html)
        self.assertIsNotNone(out)
        self.assertEqual(out.get("generated_at"), "2026-01-01")
        self.assertEqual(len(out.get("rows", [])), 1)
        self.assertEqual(out["rows"][0]["id"], 1)

    def test_extract_report_data_nested_braces(self):
        html = '<script>const reportData = {"a": 1, "b": {"c": 2}};</script>'
        out = _extract_report_data_from_html(html)
        self.assertIsNotNone(out)
        self.assertEqual(out["a"], 1)
        self.assertEqual(out["b"]["c"], 2)

    def test_extract_report_data_not_found(self):
        html = "<html><body>no reportData here</body></html>"
        out = _extract_report_data_from_html(html)
        self.assertIsNone(out)

    def test_extract_json_object_after(self):
        html = 'prefix {"x": 1, "y": 2} rest'
        out = _extract_json_object_after(html, "prefix ")
        self.assertIsNotNone(out)
        self.assertEqual(out, {"x": 1, "y": 2})


class TestFilterByDate(unittest.TestCase):
    def test_filter_rows_by_date(self):
        payload = {
            "rows": [
                {"jira_start_date": "2026-01-15", "jira_due_date": "2026-01-20"},
                {"jira_start_date": "2026-02-01", "jira_due_date": "2026-02-10"},
                {"jira_start_date": "2025-12-01", "jira_due_date": "2025-12-31"},
            ]
        }
        out = _filter_rows_by_date(payload, "2026-01-01", "2026-01-31")
        self.assertEqual(len(out["rows"]), 1)
        self.assertEqual(out["rows"][0]["jira_start_date"], "2026-01-15")


class TestInjectData(unittest.TestCase):
    def test_inject_embedded_data(self):
        html = '<html><script>const reportData = {"old": 1};</script></html>'
        new_data = {"new": 2, "rows": []}
        result = _inject_embedded_data_into_html(html, new_data)
        self.assertIn('"new": 2', result)
        self.assertIn('"rows": []', result)
        self.assertNotIn('"old": 1', result)

    def test_fetch_override_script_contains_bundle(self):
        bundle = {"GET /api/test": {"ok": True}}
        script = _make_fetch_override_script(bundle)
        self.assertIn("OFFLINE_API_BUNDLE", script)
        self.assertIn("fetch", script)
        self.assertIn('"ok": true', script)


class TestRunPrepareJob(unittest.TestCase):
    def test_run_prepare_job_embedded_report(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            base = Path(td)
            (base / "report_html").mkdir(parents=True, exist_ok=True)
            source_html = base / "report_html" / "missed_entries.html"
            source_html.write_text(
                '<!DOCTYPE html><html><body><script>const reportData = {"rows": [{"issue_key": "X-1", "jira_start_date": "2026-02-01", "jira_due_date": "2026-02-28"}], "generated_at": "test"};</script></body></html>',
                encoding="utf-8",
            )

            def resolve_sources(basedir: Path):
                return {"missed_entries.html": basedir / "report_html" / "missed_entries.html"}

            def sync_assets(_b: Path, _t: Path):
                pass

            def fetch_api(_m: str, _p: str, _q: dict | None):
                return None

            output_dir, errors = run_prepare_job(
                base_dir=base,
                from_date="2026-02-01",
                to_date="2026-02-28",
                report_keys=["missed_entries"],
                resolve_sources=resolve_sources,
                sync_assets=sync_assets,
                fetch_api=fetch_api,
            )
            self.assertTrue(output_dir.exists())
            self.assertTrue(output_dir.is_dir())
            out_html = output_dir / "missed_entries.html"
            out_json = output_dir / "missed_entries.json"
            self.assertTrue(out_html.exists(), f"Expected {out_html} to exist")
            self.assertTrue(out_json.exists(), f"Expected {out_json} to exist")
            html_content = out_html.read_text(encoding="utf-8")
            self.assertIn("reportData", html_content)
            self.assertIn("X-1", html_content)
            data = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertIn("rows", data)
            self.assertEqual(len(data["rows"]), 1)
            self.assertEqual(data["rows"][0]["issue_key"], "X-1")

    def test_run_prepare_job_skips_unknown_report(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            base = Path(td)
            (base / "report_html").mkdir(parents=True, exist_ok=True)

            def resolve_sources(basedir: Path):
                return {}

            def sync_assets(_b: Path, _t: Path):
                pass

            def fetch_api(_m: str, _p: str, _q: dict | None):
                return None

            output_dir, errors = run_prepare_job(
                base_dir=base,
                from_date="2026-01-01",
                to_date="2026-01-31",
                report_keys=["unknown_report"],
                resolve_sources=resolve_sources,
                sync_assets=sync_assets,
                fetch_api=fetch_api,
            )
            self.assertTrue(output_dir.exists())
            self.assertGreater(len(errors), 0)
            self.assertIn("Unknown", errors[0])


class TestCreateZip(unittest.TestCase):
    def test_create_zip_of_folder(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            folder = Path(td) / "sub"
            folder.mkdir(parents=True)
            (folder / "a.txt").write_text("hello", encoding="utf-8")
            (folder / "b.txt").write_text("world", encoding="utf-8")
            zip_path = create_zip_of_folder(folder)
            self.assertTrue(zip_path.exists())
            self.assertEqual(zip_path.suffix, ".zip")
            self.assertEqual(zip_path.parent, folder.parent)
            self.assertEqual(zip_path.name, "sub.zip")
            import zipfile as zf
            with zf.ZipFile(zip_path, "r") as z:
                names = sorted(z.namelist())
                self.assertEqual(names, ["a.txt", "b.txt"])


if __name__ == "__main__":
    unittest.main()
