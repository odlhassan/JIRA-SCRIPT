from pathlib import Path
import tempfile
import unittest

from azure_deploy_storage import prune_inactive_packages


class AzureDeployStorageTests(unittest.TestCase):
    def test_removes_only_inactive_zips(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "packagename.txt").write_text("active.zip")
            for name in ("active.zip", "old.zip", "customer.db", "backup.bak"):
                (root / name).write_bytes(b"preserve customer data")
            (root / "backups").mkdir()
            (root / "backups" / "saved.zip").write_bytes(b"keep backup")
            result = prune_inactive_packages(root)
            self.assertEqual([r["name"] for r in result["removed"]], ["old.zip"])
            self.assertEqual({p.name for p in root.iterdir()}, {
                "active.zip", "packagename.txt", "customer.db", "backup.bak", "backups"})
            self.assertTrue((root / "backups" / "saved.zip").exists())

    def test_missing_or_invalid_active_package_aborts_without_deletion(self):
        for active in ("missing.zip", "../outside.zip", ""):
            with self.subTest(active=active), tempfile.TemporaryDirectory() as td:
                root = Path(td)
                (root / "packagename.txt").write_text(active)
                (root / "old.zip").write_bytes(b"keep until active validated")
                with self.assertRaises(ValueError):
                    prune_inactive_packages(root)
                self.assertTrue((root / "old.zip").exists())


if __name__ == "__main__":
    unittest.main()
