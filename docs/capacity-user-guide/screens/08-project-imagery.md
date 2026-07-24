# Project Imagery (Thumbnail + Logo)

Upload a **thumbnail** (square) and a **logo** (horizontal) per project in Project Settings.
The images appear on the **RnD Muscle Utilization** product-wise view — the thumbnail in the
project tab strip (tight space) and the logo in each product panel header (wide space).

## Business Logic

- The server requires **Pillow** (installed from root `requirements.txt`) to decode, validate,
  resize, and re-encode uploads. The Azure deployment vendors this dependency into
  `.python_packages/lib/site-packages/PIL` and verifies that directory before creating the ZIP.
- Each managed project may have two images: a **thumbnail** (square) and a **logo** (horizontal).
- Uploaded files are validated (real image decode, ≤ 2 MB) and re-encoded on the server:
  - Thumbnail → center-cover cropped to a square and resized to **256×256**.
  - Logo → scaled to fit within **512×128** preserving aspect ratio; must read wider than tall
    (aspect ≥ 1.2) or it is rejected with a clear message.
  - Output format is **WEBP** for opaque art, **PNG** when the source has transparency (so
    transparent logos never get a solid box behind them).
- **Auto-thumbnail:** if a project has a logo but no explicit thumbnail, the system derives one by
  center-square-cropping the logo (flagged `thumbnail_is_auto = 1`). An explicitly uploaded
  thumbnail always wins. Removing an explicit thumbnail reverts to the auto one; removing the logo
  also removes any auto-derived thumbnail.
- Filenames are content-addressed (`<KEY>__<variant>__<sha16>.<ext>`) so a replaced image gets a
  new URL and browsers cache each version immutably.
- When no image is set, the RnD view falls back to the existing color + initials chip / display
  name — no regression for un-imaged projects.

## Business Cases

- Product/brand recognition on the RnD Muscle Utilization product-wise view, where planners scan
  many project panels quickly.
- Consistent visual identity across tight (tabs) and wide (headers) layouts without manual resizing.

## Examples

- Upload a 600×150 PNG logo for `EPR` → stored as `EPR__logo__<hash>.webp` (512×128), and an auto
  256×256 thumbnail is generated from its center square.
- Upload a 400×400 transparent PNG thumbnail for `EPR` → stored as `EPR__thumbnail__<hash>.png`
  (256×256, transparency preserved), replacing the auto one.
- Upload a 150×600 portrait image as a logo → rejected: "Logo should be wider than tall".

## Explanations

An operator opens **Project Settings → Modify** on a project, drops (or clicks to choose) images
into the two shape-matched drop zones, and sees a live "Where this appears" preview of the tab and
header. The server normalizes the image, saves the file to the persistent image directory, records
metadata in the standalone `project_images.db`, and returns the new URL. The RnD Muscle Utilization
view reads these URLs (joined by `project_key`) and renders them, falling back to color/initials
when absent.

## Front-end UI Fields

| Field / Control | Type | Default | What it controls | Example |
|---|---|---|---|---|
| Thumbnail drop zone | drag/drop + file picker | empty | Uploads/replaces the square thumbnail | Drop `icon.png` |
| Logo drop zone | drag/drop + file picker | empty | Uploads/replaces the horizontal logo | Drop `wordmark.png` |
| Replace (`swap_horiz`) | icon button (hover toolbar) | — | Re-open picker for that slot | — |
| Remove (`delete`) | icon button (hover toolbar) | — | Delete that image | — |
| "Auto from logo" badge | badge | shown when thumbnail is auto | Indicates the thumbnail was cropped from the logo | — |
| "Where this appears" strip | live preview | reflects current name/color/images | Shows mock tab + header before saving | — |
| Accepted types | — | PNG, JPG, WEBP, GIF, BMP | Input formats (SVG not accepted in v1) | — |
| Max size | — | 2 MB | Hard upload cap | — |

## Script Files

| File | Role |
|---|---|
| `project_image_registry.py` | Standalone DB init, Pillow normalization, auto-thumbnail, set/clear/get helpers, path resolution. |
| `report_server.py` | Upload/delete/serve routes, `/api/projects` image URLs, Project Settings drop-zone UI, RnD view rendering (tabs + panel logo). |
| `requirements.txt` | Declares Pillow as a production runtime dependency. |
| `.github/workflows/azure-appservice-deploy.yml` | Vendors dependencies and verifies that the `PIL` package is present before ZIP deployment. |
| `migrations/2026-07-23_project_images.py` | Idempotent creation of `project_images.db`. |
| `tests/test_project_images_api.py` | Unit + API tests. |
| `tests/test_azure_deploy_contract.py` | Guards the Pillow requirement and deployment-package verification. |
| `db_schema_changelog.py` | Records the `ADD_TABLE` change and schema snapshot. |

## Change Notes

- **2026-07-24:** Added Pillow to the production dependency set and made the Azure workflow
  verify the vendored `PIL` package before deployment. This prevents logo and thumbnail uploads
  from reaching production without the image decoder required by `normalize_image_bytes`.

## Dependent & Impacted Files

- **RnD Muscle Utilization** (`rnd_muscle_utilization_service.py`, RnD view HTML/JS in
  `report_server.py`) — consumes `thumbnail_url` / `logo_url` per `project_key`. See
  [16-rnd-muscle-utilization-settings.md](../../report-user-guide/screens/16-rnd-muscle-utilization-settings.md).
- **Managed Projects** (`managed_projects_registry.py`) — provides `project_key` / `display_name` /
  `color_hex` used for joins and fallbacks. Not modified (isolation preserved).

## Table Schema

Standalone DB `project_images.db`, table `project_images`:

| Column | Type | Meaning |
|---|---|---|
| `project_key` | TEXT PK | Managed project key. |
| `thumbnail_path`, `thumbnail_mime`, `thumbnail_width`, `thumbnail_height`, `thumbnail_sha256` | TEXT/INT | Stored square thumbnail file + metadata. |
| `thumbnail_is_auto` | INTEGER | 1 = auto-cropped from logo; 0 = explicitly uploaded. |
| `thumbnail_updated_at_utc` | TEXT | Last thumbnail change. |
| `logo_path`, `logo_mime`, `logo_width`, `logo_height`, `logo_sha256` | TEXT/INT | Stored logo file + metadata. |
| `logo_updated_at_utc` | TEXT | Last logo change. |
| `created_at_utc`, `updated_at_utc` | TEXT | Row timestamps. |

## Data Flow

1. Operator drops a file → `POST /api/projects/<key>/image/<variant>` (multipart).
2. `set_project_image` validates + normalizes (Pillow), writes the file atomically under
   `JIRA_PROJECT_IMAGE_DIR`, upserts metadata, deletes the superseded file, and derives an auto
   thumbnail when appropriate.
3. Response returns `{thumbnail_url, logo_url, dimensions, thumbnail_is_auto}`; the UI re-renders.
4. `GET /api/projects` merges image URLs into each project row.
5. RnD Muscle Utilization state (`/api/rnd-muscle-utilization`) merges image URLs into each
   `project_tab`; the view renders thumbnails in tabs and logos in product panel headers.
6. `GET /project-images/<file>` serves the bytes with `Cache-Control: immutable`.
