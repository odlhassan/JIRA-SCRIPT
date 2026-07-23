"""Standalone registry for per-project imagery (thumbnail + logo).

Kept deliberately isolated from the large ``assignee_hours_capacity.db`` so this
feature never forces a migration/download/re-upload of that multi-gigabyte file.
Image metadata lives in its own small ``project_images.db`` and the normalized
image bytes live as files under the same persistent directory.
"""

from __future__ import annotations

import hashlib
import io
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

_PROJECT_KEY_PATTERN = re.compile(r"^[A-Z0-9_-]+$")

VARIANTS: tuple[str, ...] = ("thumbnail", "logo")

MAX_UPLOAD_BYTES = 2 * 1024 * 1024  # 2 MB hard cap per file

THUMBNAIL_SIZE = 256
LOGO_MAX_HEIGHT = 128
LOGO_MAX_WIDTH = 512
LOGO_MIN_ASPECT = 1.2  # width / height; logos should read wider than tall

# Extensions Pillow can decode that we accept as input.
_ACCEPTED_INPUT_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def normalize_project_key(value: object) -> str:
    key = ("" if value is None else str(value)).strip().upper()
    if not key:
        raise ValueError("project_key is required.")
    if not _PROJECT_KEY_PATTERN.match(key):
        raise ValueError("project_key must match ^[A-Z0-9_-]+$.")
    return key


def normalize_variant(value: object) -> str:
    variant = ("" if value is None else str(value)).strip().lower()
    if variant not in VARIANTS:
        raise ValueError("variant must be 'thumbnail' or 'logo'.")
    return variant


def resolve_project_image_paths(base_dir: Path) -> dict[str, Path]:
    """Resolve a writable directory for image files + the metadata DB.

    Order of preference:
    1. ``JIRA_PROJECT_IMAGE_DIR`` env var (absolute, or relative to base_dir).
    2. ``<base_dir>/data/project_images``.
    3. ``$HOME/data/project_images`` (Azure App Service persistent share).
    """
    base_dir = Path(base_dir)
    raw = os.getenv("JIRA_PROJECT_IMAGE_DIR", "").strip().strip('"').strip("'")
    azure_home = os.getenv("HOME", "").strip()
    on_azure = bool(os.getenv("WEBSITE_INSTANCE_ID") or os.getenv("WEBSITE_SITE_NAME"))
    if raw:
        images_dir = Path(raw)
        if not images_dir.is_absolute():
            images_dir = base_dir / images_dir
    elif on_azure and azure_home:
        # wwwroot is writable but wiped on redeploy; /home persists across deploys.
        images_dir = Path(azure_home) / "data" / "project_images"
    else:
        images_dir = base_dir / "data" / "project_images"

    if not _dir_is_writable(images_dir):
        azure_home = os.getenv("HOME", "")
        fallback = (
            Path(azure_home) / "data" / "project_images"
            if azure_home
            else Path("data") / "project_images"
        )
        images_dir = fallback
        images_dir.mkdir(parents=True, exist_ok=True)

    return {"images_dir": images_dir, "db_path": images_dir / "project_images.db"}


def _dir_is_writable(candidate: Path) -> bool:
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        probe = candidate / ".write-probe"
        with open(probe, "a", encoding="utf-8"):
            pass
        probe.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def init_project_images_db(db_path: Path) -> None:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS project_images (
                project_key TEXT PRIMARY KEY,
                thumbnail_path TEXT NOT NULL DEFAULT '',
                thumbnail_mime TEXT NOT NULL DEFAULT '',
                thumbnail_width INTEGER NOT NULL DEFAULT 0,
                thumbnail_height INTEGER NOT NULL DEFAULT 0,
                thumbnail_sha256 TEXT NOT NULL DEFAULT '',
                thumbnail_is_auto INTEGER NOT NULL DEFAULT 0,
                thumbnail_updated_at_utc TEXT NOT NULL DEFAULT '',
                logo_path TEXT NOT NULL DEFAULT '',
                logo_mime TEXT NOT NULL DEFAULT '',
                logo_width INTEGER NOT NULL DEFAULT 0,
                logo_height INTEGER NOT NULL DEFAULT 0,
                logo_sha256 TEXT NOT NULL DEFAULT '',
                logo_updated_at_utc TEXT NOT NULL DEFAULT '',
                created_at_utc TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _row_to_dict(row: sqlite3.Row | None) -> dict:
    if row is None:
        return {}
    return {
        "project_key": str(row["project_key"] or ""),
        "thumbnail_path": str(row["thumbnail_path"] or ""),
        "thumbnail_mime": str(row["thumbnail_mime"] or ""),
        "thumbnail_width": int(row["thumbnail_width"] or 0),
        "thumbnail_height": int(row["thumbnail_height"] or 0),
        "thumbnail_sha256": str(row["thumbnail_sha256"] or ""),
        "thumbnail_is_auto": bool(int(row["thumbnail_is_auto"] or 0)),
        "thumbnail_updated_at_utc": str(row["thumbnail_updated_at_utc"] or ""),
        "logo_path": str(row["logo_path"] or ""),
        "logo_mime": str(row["logo_mime"] or ""),
        "logo_width": int(row["logo_width"] or 0),
        "logo_height": int(row["logo_height"] or 0),
        "logo_sha256": str(row["logo_sha256"] or ""),
        "logo_updated_at_utc": str(row["logo_updated_at_utc"] or ""),
    }


def get_project_image_record(db_path: Path, project_key: str) -> dict:
    key = normalize_project_key(project_key)
    init_project_images_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT * FROM project_images WHERE project_key = ?", (key,)
        ).fetchone()
        return _row_to_dict(row)
    finally:
        conn.close()


def get_project_image_map(db_path: Path) -> dict[str, dict]:
    """Return {project_key: record} for every project that has any image set."""
    init_project_images_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute("SELECT * FROM project_images").fetchall()
    finally:
        conn.close()
    out: dict[str, dict] = {}
    for row in rows:
        record = _row_to_dict(row)
        if record.get("thumbnail_path") or record.get("logo_path"):
            out[record["project_key"]] = record
    return out


def _has_alpha(image) -> bool:
    if image.mode in ("RGBA", "LA", "PA"):
        return True
    return image.mode == "P" and "transparency" in image.info


def normalize_image_bytes(data: bytes, variant: str) -> tuple[bytes, str, str, int, int]:
    """Validate + re-encode an uploaded image for the target variant.

    Returns ``(out_bytes, ext, mime, width, height)``. WEBP is used for opaque
    art; PNG is used when the source has transparency so logos never gain a
    solid box behind them.
    """
    from PIL import Image, ImageOps, UnidentifiedImageError

    variant = normalize_variant(variant)
    if not data:
        raise ValueError("Uploaded file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(
            f"Image is {len(data) / (1024 * 1024):.1f} MB - the maximum is 2 MB."
        )

    try:
        image = Image.open(io.BytesIO(data))
        image.load()
    except (UnidentifiedImageError, OSError) as exc:
        raise ValueError("File is not a readable image (PNG, JPG, WEBP, GIF or BMP).") from exc

    image = ImageOps.exif_transpose(image)
    src_w, src_h = image.size
    if src_w <= 0 or src_h <= 0:
        raise ValueError("Image has invalid dimensions.")

    keep_alpha = _has_alpha(image)

    if variant == "thumbnail":
        # Forgiving: center-cover crop any aspect to a clean square.
        image = ImageOps.fit(
            image, (THUMBNAIL_SIZE, THUMBNAIL_SIZE), method=Image.LANCZOS, centering=(0.5, 0.5)
        )
    else:  # logo
        if src_w / src_h < LOGO_MIN_ASPECT:
            raise ValueError(
                "Logo should be wider than tall (roughly 3:1). Use the thumbnail slot for square art."
            )
        scale = min(LOGO_MAX_WIDTH / src_w, LOGO_MAX_HEIGHT / src_h, 1.0)
        target = (max(1, round(src_w * scale)), max(1, round(src_h * scale)))
        if target != (src_w, src_h):
            image = image.resize(target, Image.LANCZOS)

    if keep_alpha:
        out_image = image.convert("RGBA")
        buffer = io.BytesIO()
        out_image.save(buffer, format="PNG", optimize=True)
        ext, mime = ".png", "image/png"
    else:
        out_image = image.convert("RGB")
        buffer = io.BytesIO()
        quality = 82 if variant == "thumbnail" else 85
        out_image.save(buffer, format="WEBP", quality=quality, method=6)
        ext, mime = ".webp", "image/webp"

    out_bytes = buffer.getvalue()
    return out_bytes, ext, mime, out_image.width, out_image.height


def _filename_for(project_key: str, variant: str, digest: str, ext: str) -> str:
    return f"{project_key}__{variant}__{digest[:16]}{ext}"


def _write_atomic(images_dir: Path, filename: str, data: bytes) -> None:
    images_dir.mkdir(parents=True, exist_ok=True)
    target = images_dir / filename
    tmp = images_dir / (filename + ".tmp")
    with open(tmp, "wb") as handle:
        handle.write(data)
    os.replace(tmp, target)


def _delete_file(images_dir: Path, filename: str) -> None:
    if not filename:
        return
    try:
        (Path(images_dir) / filename).unlink(missing_ok=True)
    except OSError:
        pass


def _upsert(conn: sqlite3.Connection, project_key: str, fields: dict) -> None:
    now = _utc_now_iso()
    existing = conn.execute(
        "SELECT project_key FROM project_images WHERE project_key = ?", (project_key,)
    ).fetchone()
    if existing is None:
        fields = {**fields, "created_at_utc": now, "updated_at_utc": now}
        cols = ", ".join(fields.keys())
        placeholders = ", ".join(["?"] * len(fields))
        conn.execute(
            f"INSERT INTO project_images (project_key, {cols}) VALUES (?, {placeholders})",
            (project_key, *fields.values()),
        )
    else:
        fields = {**fields, "updated_at_utc": now}
        assignments = ", ".join(f"{col} = ?" for col in fields.keys())
        conn.execute(
            f"UPDATE project_images SET {assignments} WHERE project_key = ?",
            (*fields.values(), project_key),
        )


def set_project_image(
    db_path: Path,
    images_dir: Path,
    project_key: str,
    variant: str,
    data: bytes,
) -> dict:
    """Store one image variant. Uploading a logo auto-derives the thumbnail
    when no explicit thumbnail exists (or the current one was auto-derived)."""
    key = normalize_project_key(project_key)
    variant = normalize_variant(variant)
    images_dir = Path(images_dir)
    init_project_images_db(db_path)

    out_bytes, ext, mime, width, height = normalize_image_bytes(data, variant)
    digest = hashlib.sha256(out_bytes).hexdigest()
    now = _utc_now_iso()

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        current = _row_to_dict(
            conn.execute(
                "SELECT * FROM project_images WHERE project_key = ?", (key,)
            ).fetchone()
        )

        if variant == "thumbnail":
            filename = _filename_for(key, "thumbnail", digest, ext)
            _write_atomic(images_dir, filename, out_bytes)
            old = current.get("thumbnail_path", "")
            _upsert(
                conn,
                key,
                {
                    "thumbnail_path": filename,
                    "thumbnail_mime": mime,
                    "thumbnail_width": width,
                    "thumbnail_height": height,
                    "thumbnail_sha256": digest,
                    "thumbnail_is_auto": 0,
                    "thumbnail_updated_at_utc": now,
                },
            )
            if old and old != filename:
                _delete_file(images_dir, old)
        else:  # logo
            filename = _filename_for(key, "logo", digest, ext)
            _write_atomic(images_dir, filename, out_bytes)
            old_logo = current.get("logo_path", "")
            _upsert(
                conn,
                key,
                {
                    "logo_path": filename,
                    "logo_mime": mime,
                    "logo_width": width,
                    "logo_height": height,
                    "logo_sha256": digest,
                    "logo_updated_at_utc": now,
                },
            )
            if old_logo and old_logo != filename:
                _delete_file(images_dir, old_logo)

            # Auto-derive a thumbnail from the logo when none is explicitly set.
            if not current.get("thumbnail_path") or current.get("thumbnail_is_auto"):
                _write_auto_thumbnail_from_bytes(conn, images_dir, key, out_bytes, current)

        conn.commit()
        row = conn.execute(
            "SELECT * FROM project_images WHERE project_key = ?", (key,)
        ).fetchone()
        return _row_to_dict(row)
    finally:
        conn.close()


def _write_auto_thumbnail_from_bytes(
    conn: sqlite3.Connection,
    images_dir: Path,
    project_key: str,
    logo_bytes: bytes,
    current: dict,
) -> None:
    thumb_bytes, ext, mime, width, height = normalize_image_bytes(logo_bytes, "thumbnail")
    digest = hashlib.sha256(thumb_bytes).hexdigest()
    filename = _filename_for(project_key, "thumbnail", digest, ext)
    _write_atomic(images_dir, filename, thumb_bytes)
    old = current.get("thumbnail_path", "")
    _upsert(
        conn,
        project_key,
        {
            "thumbnail_path": filename,
            "thumbnail_mime": mime,
            "thumbnail_width": width,
            "thumbnail_height": height,
            "thumbnail_sha256": digest,
            "thumbnail_is_auto": 1,
            "thumbnail_updated_at_utc": _utc_now_iso(),
        },
    )
    if old and old != filename:
        _delete_file(images_dir, old)


def clear_project_image(
    db_path: Path,
    images_dir: Path,
    project_key: str,
    variant: str,
) -> dict:
    """Remove one image variant.

    Removing an explicit thumbnail re-derives an auto thumbnail from the logo
    if a logo exists. Removing a logo also removes any auto-derived thumbnail.
    """
    key = normalize_project_key(project_key)
    variant = normalize_variant(variant)
    images_dir = Path(images_dir)
    init_project_images_db(db_path)

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        current = _row_to_dict(
            conn.execute(
                "SELECT * FROM project_images WHERE project_key = ?", (key,)
            ).fetchone()
        )
        if not current:
            return {}

        if variant == "thumbnail":
            _delete_file(images_dir, current.get("thumbnail_path", ""))
            _clear_thumbnail_fields(conn, key)
            if current.get("logo_path"):
                logo_file = images_dir / current["logo_path"]
                try:
                    logo_bytes = logo_file.read_bytes()
                except OSError:
                    logo_bytes = b""
                if logo_bytes:
                    fresh = _row_to_dict(
                        conn.execute(
                            "SELECT * FROM project_images WHERE project_key = ?", (key,)
                        ).fetchone()
                    )
                    _write_auto_thumbnail_from_bytes(conn, images_dir, key, logo_bytes, fresh)
        else:  # logo
            _delete_file(images_dir, current.get("logo_path", ""))
            _clear_logo_fields(conn, key)
            if current.get("thumbnail_is_auto"):
                _delete_file(images_dir, current.get("thumbnail_path", ""))
                _clear_thumbnail_fields(conn, key)

        conn.commit()
        row = conn.execute(
            "SELECT * FROM project_images WHERE project_key = ?", (key,)
        ).fetchone()
        return _row_to_dict(row)
    finally:
        conn.close()


def _clear_thumbnail_fields(conn: sqlite3.Connection, project_key: str) -> None:
    conn.execute(
        """
        UPDATE project_images
        SET thumbnail_path = '', thumbnail_mime = '', thumbnail_width = 0,
            thumbnail_height = 0, thumbnail_sha256 = '', thumbnail_is_auto = 0,
            thumbnail_updated_at_utc = '', updated_at_utc = ?
        WHERE project_key = ?
        """,
        (_utc_now_iso(), project_key),
    )


def _clear_logo_fields(conn: sqlite3.Connection, project_key: str) -> None:
    conn.execute(
        """
        UPDATE project_images
        SET logo_path = '', logo_mime = '', logo_width = 0, logo_height = 0,
            logo_sha256 = '', logo_updated_at_utc = '', updated_at_utc = ?
        WHERE project_key = ?
        """,
        (_utc_now_iso(), project_key),
    )
