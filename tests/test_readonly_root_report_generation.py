from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

import report_server
from generate_rlt_leave_report import _write_xlsx


class ResolveScriptCwdTests(unittest.TestCase):
    def test_writable_root_is_used_unchanged(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            self.assertEqual(report_server._resolve_script_cwd(root), root)

    def test_explicit_cwd_wins(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            override = root / "elsewhere"
            self.assertEqual(report_server._resolve_script_cwd(root, override), override)

    def test_read_only_root_falls_back_to_writable_artifact_dir(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            artifacts = root / "artifacts"
            original = report_server._is_writable_directory
            report_server._is_writable_directory = lambda path: False
            os.environ["JIRA_CANONICAL_ARTIFACT_DIR"] = str(artifacts)
            try:
                resolved = report_server._resolve_script_cwd(root / "wwwroot")
            finally:
                report_server._is_writable_directory = original
                os.environ.pop("JIRA_CANONICAL_ARTIFACT_DIR", None)
            self.assertEqual(resolved, artifacts)

    def test_run_script_uses_resolved_cwd(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            work = root / "work"
            work.mkdir()
            (root / "probe_cwd.py").write_text("import os; print(os.getcwd())", encoding="utf-8")
            code, stdout, _stderr = report_server._run_script("probe_cwd.py", root, cwd=work)
            self.assertEqual(code, 0)
            self.assertEqual(Path(stdout.strip()).resolve(), work.resolve())


class WriteXlsxReadOnlyFallbackTests(unittest.TestCase):
    def _empty_aggregates(self) -> dict:
        return {
            "assignee_summary": [],
            "daily": [],
            "weekly": [],
            "monthly": [],
            "defective": [],
            "clubbed": [],
        }

    def _patch_readonly_mkstemp(self, gen):
        real_mkstemp = gen.tempfile.mkstemp

        def _fake_mkstemp(*args, **kwargs):
            if kwargs.get("dir"):
                raise OSError(30, "Read-only file system")
            return real_mkstemp(*args, **kwargs)

        gen.tempfile.mkstemp = _fake_mkstemp
        self.addCleanup(setattr, gen.tempfile, "mkstemp", real_mkstemp)

    def test_read_only_target_dir_still_writes_workbook(self):
        import generate_rlt_leave_report as gen

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            target = Path(td) / "out" / "rlt_leave_report.xlsx"
            self._patch_readonly_mkstemp(gen)
            _write_xlsx(target, [], [], [], self._empty_aggregates())
            self.assertTrue(target.exists())
            self.assertGreater(target.stat().st_size, 0)

    def test_unreplaceable_target_warns_instead_of_raising(self):
        import generate_rlt_leave_report as gen

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            target = Path(td) / "out" / "rlt_leave_report.xlsx"
            self._patch_readonly_mkstemp(gen)
            real_replace = gen.os.replace

            def _fake_replace(src, dst):
                raise OSError(30, "Read-only file system")

            gen.os.replace = _fake_replace
            self.addCleanup(setattr, gen.os, "replace", real_replace)
            with self.assertWarns(UserWarning):
                _write_xlsx(target, [], [], [], self._empty_aggregates())


if __name__ == "__main__":
    unittest.main()
