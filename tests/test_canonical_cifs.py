"""Linux allocation safety checks, using a synthetic CIFS server response."""
import os
from pathlib import Path
import sqlite3
import struct
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import canonical_storage


@unittest.skipUnless(sys.platform == "linux", "CIFS maintenance runs on Linux")
class CifsAllocationSafetyTests(unittest.TestCase):
    def exercise(self, failed):
        with tempfile.TemporaryDirectory() as folder:
            db = Path(folder) / "scratch.db"
            with sqlite3.connect(db) as conn:
                conn.execute("CREATE TABLE saved(value)")
                conn.execute("INSERT INTO saved VALUES ('keep')")
            original = db.read_bytes()
            responses = iter([(4194304, len(original)), (65536, len(original))])

            def query(fd, code, buf, mutate):
                allocation, size = next(responses)
                struct.pack_into("<QQ", buf, 24, allocation, size)

            libc = MagicMock()
            def punch(fd, mode, start, length):
                self.assertEqual(os.fstat(fd).st_size, 4194304)
                self.assertGreaterEqual(start, len(original))
                return -1 if failed else 0
            libc.fallocate.side_effect = punch
            with patch("fcntl.ioctl", side_effect=query), patch("ctypes.CDLL", return_value=libc), patch("ctypes.get_errno", return_value=5):
                if failed:
                    with self.assertRaises(OSError):
                        canonical_storage._trim_cifs_in_process(db)
                else:
                    result = canonical_storage._trim_cifs_in_process(db)
                    self.assertEqual(result["allocated_bytes_after"], 65536)
            self.assertEqual(db.read_bytes(), original)
            with sqlite3.connect(db, timeout=0) as conn:
                conn.execute("BEGIN EXCLUSIVE")
                self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")
                self.assertEqual(conn.execute("SELECT value FROM saved").fetchone()[0], "keep")

    def test_success_preserves_prefix_eof_and_releases_lock(self):
        self.exercise(False)

    def test_punch_failure_restores_eof_and_releases_lock(self):
        self.exercise(True)


if __name__ == "__main__":
    unittest.main()
