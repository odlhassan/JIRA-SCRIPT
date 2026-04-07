from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import run_server


class RunServerTests(unittest.TestCase):
    def test_main_rebuilds_html_reports_before_serving_by_default(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            with patch.object(run_server, "__file__", str(base_dir / "run_server.py")):
                with patch("sys.argv", ["run_server.py"]):
                    with patch.object(run_server, "_clear_startup_caches") as clear_cache_mock:
                        with patch.object(run_server, "rebuild_html_reports") as rebuild_mock:
                            with patch.object(run_server, "_resolve_server_port", return_value=3000):
                                with patch.object(run_server, "run_report_server") as serve_mock:
                                    run_server.main()

        clear_cache_mock.assert_called_once_with(base_dir)
        rebuild_mock.assert_called_once_with(
            base_dir,
            "report_html",
            include_dashboard=False,
            skip_phase_rmi_gantt=False,
            skip_ipp_dashboard=False,
        )
        serve_mock.assert_called_once_with(
            base_dir=base_dir,
            folder_raw="report_html",
            host="127.0.0.1",
            port=3000,
        )

    def test_main_skips_startup_rebuild_when_no_sync_is_requested(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            with patch.object(run_server, "__file__", str(base_dir / "run_server.py")):
                with patch("sys.argv", ["run_server.py", "--no-sync"]):
                    with patch.object(run_server, "_clear_startup_caches") as clear_cache_mock:
                        with patch.object(run_server, "rebuild_html_reports") as rebuild_mock:
                            with patch.object(run_server, "_resolve_server_port", return_value=3000):
                                with patch.object(run_server, "run_report_server") as serve_mock:
                                    run_server.main()

        clear_cache_mock.assert_called_once_with(base_dir)
        rebuild_mock.assert_not_called()
        serve_mock.assert_called_once_with(
            base_dir=base_dir,
            folder_raw="report_html",
            host="127.0.0.1",
            port=3000,
        )

    def test_main_forwards_rebuild_flags_and_default_fresh_cache_clear(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            argv = [
                "run_server.py",
                "--fresh",
                "--include-dashboard",
                "--skip-phase-rmi-gantt",
                "--skip-ipp-dashboard",
                "--report-html-dir",
                "custom_reports",
                "--host",
                "0.0.0.0",
                "--port",
                "8123",
            ]
            with patch.object(run_server, "__file__", str(base_dir / "run_server.py")):
                with patch("sys.argv", argv):
                    with patch.object(run_server, "_clear_startup_caches") as clear_cache_mock:
                        with patch.object(run_server, "rebuild_html_reports") as rebuild_mock:
                            with patch.object(run_server, "_resolve_server_port", return_value=8123) as resolve_port_mock:
                                with patch.object(run_server, "run_report_server") as serve_mock:
                                    run_server.main()

        clear_cache_mock.assert_called_once_with(base_dir)
        rebuild_mock.assert_called_once_with(
            base_dir,
            "custom_reports",
            include_dashboard=True,
            skip_phase_rmi_gantt=True,
            skip_ipp_dashboard=True,
        )
        resolve_port_mock.assert_called_once_with("0.0.0.0", 8123)
        serve_mock.assert_called_once_with(
            base_dir=base_dir,
            folder_raw="custom_reports",
            host="0.0.0.0",
            port=8123,
        )

    def test_main_skips_cache_clear_when_keep_cache_is_requested(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            with patch.object(run_server, "__file__", str(base_dir / "run_server.py")):
                with patch("sys.argv", ["run_server.py", "--keep-cache"]):
                    with patch.object(run_server, "_clear_startup_caches") as clear_cache_mock:
                        with patch.object(run_server, "rebuild_html_reports") as rebuild_mock:
                            with patch.object(run_server, "_resolve_server_port", return_value=3000):
                                with patch.object(run_server, "run_report_server") as serve_mock:
                                    run_server.main()

        clear_cache_mock.assert_not_called()
        rebuild_mock.assert_called_once()
        serve_mock.assert_called_once()


if __name__ == "__main__":
    unittest.main()
