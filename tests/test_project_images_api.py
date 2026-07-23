from __future__ import annotations

import io
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from PIL import Image

from report_server import create_report_server_app
import project_image_registry as pir


def _png_bytes(width: int, height: int, *, alpha: bool = False, color=(200, 30, 30)) -> bytes:
    mode = "RGBA" if alpha else "RGB"
    fill = color + ((255,) if alpha else ())
    image = Image.new(mode, (width, height), fill)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


class ProjectImageRegistryTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.dir = Path(self._tmp.name)
        self.db = self.dir / "project_images.db"
        pir.init_project_images_db(self.db)

    def tearDown(self):
        self._tmp.cleanup()

    def test_logo_upload_derives_auto_thumbnail(self):
        rec = pir.set_project_image(self.db, self.dir, "epr", "logo", _png_bytes(600, 150))
        self.assertTrue(rec["logo_path"])
        self.assertEqual(rec["logo_mime"], "image/webp")
        self.assertTrue(rec["thumbnail_path"])
        self.assertTrue(rec["thumbnail_is_auto"])
        self.assertEqual(rec["thumbnail_width"], 256)

    def test_explicit_thumbnail_wins_and_keeps_alpha_as_png(self):
        pir.set_project_image(self.db, self.dir, "EPR", "logo", _png_bytes(600, 150))
        rec = pir.set_project_image(self.db, self.dir, "EPR", "thumbnail", _png_bytes(400, 400, alpha=True))
        self.assertFalse(rec["thumbnail_is_auto"])
        self.assertEqual(rec["thumbnail_mime"], "image/png")

    def test_clear_thumbnail_reverts_to_auto_from_logo(self):
        pir.set_project_image(self.db, self.dir, "EPR", "logo", _png_bytes(600, 150))
        pir.set_project_image(self.db, self.dir, "EPR", "thumbnail", _png_bytes(400, 400))
        rec = pir.clear_project_image(self.db, self.dir, "EPR", "thumbnail")
        self.assertTrue(rec["thumbnail_is_auto"])
        self.assertTrue(rec["thumbnail_path"])

    def test_clear_logo_removes_auto_thumbnail(self):
        pir.set_project_image(self.db, self.dir, "EPR", "logo", _png_bytes(600, 150))
        rec = pir.clear_project_image(self.db, self.dir, "EPR", "logo")
        self.assertEqual(rec.get("logo_path", ""), "")
        self.assertEqual(rec.get("thumbnail_path", ""), "")
        leftover = [p.name for p in self.dir.iterdir() if p.suffix in {".webp", ".png"}]
        self.assertEqual(leftover, [])

    def test_portrait_logo_rejected(self):
        with self.assertRaises(ValueError):
            pir.set_project_image(self.db, self.dir, "EPR", "logo", _png_bytes(150, 600))

    def test_oversize_rejected(self):
        with self.assertRaises(ValueError):
            pir.normalize_image_bytes(b"x" * (pir.MAX_UPLOAD_BYTES + 1), "thumbnail")

    def test_non_image_rejected(self):
        with self.assertRaises(ValueError):
            pir.normalize_image_bytes(b"not an image", "thumbnail")


class ProjectImageApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_upload_serve_and_delete_via_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            created = client.post(
                "/api/projects",
                json={"project_key": "PIMG", "project_name": "Pic Project", "display_name": "Pic Project", "color_hex": "#1D4ED8"},
            )
            self.assertEqual(created.status_code, 200)

            upload = client.post(
                "/api/projects/PIMG/image/logo",
                data={"file": (io.BytesIO(_png_bytes(600, 150)), "logo.png")},
                content_type="multipart/form-data",
            )
            self.assertEqual(upload.status_code, 200)
            body = upload.get_json()
            self.assertTrue(body["logo_url"].startswith("/project-images/"))
            self.assertTrue(body["thumbnail_url"])  # auto-derived
            self.assertTrue(body["thumbnail_is_auto"])

            served = client.get(body["logo_url"])
            self.assertEqual(served.status_code, 200)
            self.assertIn("immutable", served.headers.get("Cache-Control", ""))

            listing = client.get("/api/projects").get_json()["projects"]
            row = next(item for item in listing if item["project_key"] == "PIMG")
            self.assertTrue(row["logo_url"])
            self.assertTrue(row["thumbnail_url"])

            removed = client.delete("/api/projects/PIMG/image/logo")
            self.assertEqual(removed.status_code, 200)
            self.assertIsNone(removed.get_json()["logo_url"])

    def test_upload_rejects_bad_image(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()
            client.post(
                "/api/projects",
                json={"project_key": "PBAD", "project_name": "Bad", "display_name": "Bad", "color_hex": "#1D4ED8"},
            )
            resp = client.post(
                "/api/projects/PBAD/image/thumbnail",
                data={"file": (io.BytesIO(b"nope"), "x.png")},
                content_type="multipart/form-data",
            )
            self.assertEqual(resp.status_code, 400)

    def test_serve_rejects_path_traversal(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()
            resp = client.get("/project-images/..%2f..%2fassignee_hours_report.xlsx")
            self.assertEqual(resp.status_code, 404)


if __name__ == "__main__":
    unittest.main()
