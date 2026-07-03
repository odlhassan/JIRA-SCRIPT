import importlib
import os
import sqlite3
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import db_journal_mode


class DbJournalModeTests(unittest.TestCase):
    _ENV_VARS = ("WEBSITE_INSTANCE_ID", "WEBSITE_SITE_NAME", "EPR_FORCE_WAL", "EPR_DISABLE_WAL")

    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in self._ENV_VARS}
        for k in self._ENV_VARS:
            os.environ.pop(k, None)
        importlib.reload(db_journal_mode)

    def tearDown(self):
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        importlib.reload(db_journal_mode)

    def _applied_mode(self, tmp_path: Path) -> str:
        importlib.reload(db_journal_mode)
        db = tmp_path / "jm.db"
        conn = sqlite3.connect(db)
        try:
            db_journal_mode.apply_journal_mode(conn)
            return conn.execute("PRAGMA journal_mode").fetchone()[0].lower()
        finally:
            conn.close()

    def test_default_uses_rollback_journal(self):
        self.assertEqual(db_journal_mode.safe_journal_mode(), "DELETE")

    def test_azure_instance_uses_rollback_journal(self):
        os.environ["WEBSITE_INSTANCE_ID"] = "abc123"
        importlib.reload(db_journal_mode)
        self.assertEqual(db_journal_mode.safe_journal_mode(), "DELETE")

    def test_azure_site_name_uses_rollback_journal(self):
        os.environ["WEBSITE_SITE_NAME"] = "epreporting"
        importlib.reload(db_journal_mode)
        self.assertEqual(db_journal_mode.safe_journal_mode(), "DELETE")

    def test_force_wal_overrides_azure(self):
        os.environ["WEBSITE_INSTANCE_ID"] = "abc123"
        os.environ["EPR_FORCE_WAL"] = "1"
        importlib.reload(db_journal_mode)
        self.assertEqual(db_journal_mode.safe_journal_mode(), "WAL")

    def test_disable_wal_overrides_local(self):
        os.environ["EPR_DISABLE_WAL"] = "1"
        importlib.reload(db_journal_mode)
        self.assertEqual(db_journal_mode.safe_journal_mode(), "DELETE")

    def test_apply_sets_pragma_on_real_connection(self):
        import tempfile

        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(self._applied_mode(Path(d)), "delete")
            os.environ["EPR_FORCE_WAL"] = "1"
            self.assertEqual(self._applied_mode(Path(d)), "wal")

    def test_apply_does_not_kill_startup_when_journal_pragma_is_locked(self):
        conn = Mock()
        conn.execute.side_effect = sqlite3.OperationalError("database is locked")

        self.assertEqual(db_journal_mode.apply_journal_mode(conn), "DELETE")


if __name__ == "__main__":
    unittest.main()
