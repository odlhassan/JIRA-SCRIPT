# Design Document — Project Image Upload (Thumbnail + Logo)

**Status:** Implemented (2026-07-23)
**Author:** Engineering
**Date:** 2026-07-23
**Feature area:** Project Settings → RnD Muscle Utilization (product-wise view)

> **Implementation note (supersedes §3–§4 storage plan below):** to honor the DB-isolation
> preference and avoid any migration of the ~3.75 GB `assignee_hours_capacity.db`, image
> metadata is stored in a **standalone `project_images.db`** (not new columns on
> `managed_projects`). The table `project_images` is keyed by `project_key` and carries the same
> fields the columns list describes. Image bytes are files under `JIRA_PROJECT_IMAGE_DIR`
> (default `<base>/data/project_images`, `/home/data/project_images` on Azure). The standalone DB
> is auto-created at startup and by `migrations/2026-07-23_project_images.py`, so **production
> needs no download/re-upload of the large DB**. Everything else in this document is as built.

---

## 1. Summary

Allow an operator, from **Project Settings** (`/settings/projects`), to upload **two images per
project**:

1. **Thumbnail** — square (1:1), used where horizontal space is tight (tabs, chips, dense lists).
2. **Logo** — horizontal / wide (roughly 3:1–4:1), used where more space is available (headers,
   banners, the product-wise section title of RnD Muscle Utilization).

The primary consumer is the **product-wise view of RnD Muscle Utilization**
(`/settings/rnd-muscle-utilization/view`), where each product/project tab and its section header
will show the appropriate image, falling back gracefully to the existing colored initials chip when
no image is set.

---

## 2. Goals & Non-Goals

### Goals
- Upload, replace, and remove a **thumbnail** and a **logo** per project.
- Render the correct variant automatically based on available space (thumbnail in tight spots,
  logo in wide spots).
- Zero-regression fallback: projects without images look exactly as today (color + initials).
- Keep the migration on the 3.75 GB `assignee_hours_capacity.db` **cheap and fast**.

### Non-Goals
- No image editing/cropping UI in v1 (client-side validation + server-side normalization only).
- No per-user or per-role image variants.
- No CDN integration in v1 (served by the app; can be added later without schema change).

---

## 3. Key architectural decision — where the bytes live

`managed_projects` lives inside **`assignee_hours_capacity.db` (~3.75 GB)**. The user's stated
operational pain is that downloading/re-uploading this DB for a production migration is slow.

**Decision: do NOT store image bytes as BLOBs in the DB.**

- BLOBs would inflate the already-4 GB file, worsening every future download/upload migration and
  bloating SQLite page cache and backups.
- Instead, store image **files on the persistent filesystem** and keep only **small metadata
  columns** (relative path, mime, dimensions, checksum) in `managed_projects`.
- Result: the schema migration adds a handful of short `TEXT`/`INTEGER` columns — a near-instant
  `ADD_COLUMN` migration even on the 4 GB DB — and the heavy payload never touches the DB round-trip.

### Storage location (persistence-safe) — DECIDED
- **Azure App Service:** `/home/data/project_images/`. On App Service, only the `/home` share
  is **persistent and survives restarts/redeploys**; app-root content is replaced on every deploy.
  This is the confirmed production directory.
- **Local dev:** `E:\JIRA SCRIPT\data\project_images\`.
- Resolved via env var **`JIRA_PROJECT_IMAGE_DIR`**, defaulting to `/home/data/project_images`
  when running on Azure (detected via `WEBSITE_INSTANCE_ID` / `HOME` env) and to the repo
  `data/project_images/` locally. **Never** stored inside the deploy package.
- `.gitignore` must exclude `data/project_images/` so uploaded assets are never committed.

Filenames are content-addressed to bust caches and avoid collisions:
`<project_key>__<variant>__<sha256[:16]>.<ext>` (e.g. `EPR__logo__9af12c3b5e7d0a41.webp`).

---

## 4. Data model / schema change

> ⚠️ **DB structural change — full protocol applies.** Target DB: `assignee_hours_capacity.db`,
> table `managed_projects`. Follow `DATABASE_SCHEMA_MIGRATION_PROTOCOL.md` and the
> `db-schema-migration-discipline` skill: record each `ADD_COLUMN` in `db_schema_changelog.db`,
> snapshot the new schema, and prepare the `db_migration.py` rename→create→copy→drop path for the
> operator-downloaded production DB. Because these are additive nullable columns, the production
> upgrade is a fast column add (no full-table rebuild required for the bytes).

New columns on `managed_projects` (all nullable / defaulted, additive):

| Column | Type | Default | Meaning |
|---|---|---|---|
| `thumbnail_path` | TEXT | `''` | Relative path (from image dir) to the square thumbnail file. Empty = none. |
| `thumbnail_mime` | TEXT | `''` | MIME type of the stored thumbnail (e.g. `image/webp`). |
| `thumbnail_width` | INTEGER | `0` | Stored pixel width (post-normalization). |
| `thumbnail_height` | INTEGER | `0` | Stored pixel height. |
| `thumbnail_sha256` | TEXT | `''` | Content checksum for cache-busting + dedupe. |
| `thumbnail_is_auto` | INTEGER | `0` | `1` = thumbnail was auto-derived by center-cropping the logo; `0` = explicitly uploaded. |
| `thumbnail_updated_at_utc` | TEXT | `''` | Last upload/replace timestamp. |
| `logo_path` | TEXT | `''` | Relative path to the horizontal logo file. Empty = none. |
| `logo_mime` | TEXT | `''` | MIME type of the stored logo. |
| `logo_width` | INTEGER | `0` | Stored pixel width. |
| `logo_height` | INTEGER | `0` | Stored pixel height. |
| `logo_sha256` | TEXT | `''` | Content checksum. |
| `logo_updated_at_utc` | TEXT | `''` | Last upload/replace timestamp. |

Rationale for two full sets rather than a child table: exactly two, fixed image slots per project;
a 1:1 shape keeps reads join-free and the registry code (`_row_to_project`) simple.

---

## 5. Backend design

### 5.1 Files touched
| File | Change |
|---|---|
| `managed_projects_registry.py` | Add columns in `init_managed_projects_db`; extend `normalize_managed_project_payload` (ignore image fields — they're set via dedicated endpoints), `_row_to_project` to surface image metadata; add `set_project_image()` / `clear_project_image()` helpers; add image-dir resolution + file write/delete + normalization. |
| `report_server.py` | New routes (upload, delete, serve); wire image dir into `capacity_paths`; include image URLs in `/api/projects`; pass image data into RnD muscle project-tab payload; update `_projects_settings_html()` UI. |
| `rnd_muscle_utilization_service.py` | Extend `RndMuscleProjectTab` with `thumbnail_url` / `logo_url`; join `managed_projects` image metadata by `project_key` when building tabs (`list_rnd_muscle_project_tabs`, `load_rnd_muscle_utilization_page_state`). |
| `db_schema_changelog.py` (usage) | Record the 12 `ADD_COLUMN` entries + schema snapshot. |
| `db_migration.py` (usage) | Ensure additive columns are handled in the production upgrade plan. |
| `.gitignore` | Add `data/project_images/`. |

### 5.2 Routes (all under existing settings auth)
- `POST /api/projects/<project_key>/image/<variant>` — multipart upload; `variant ∈ {thumbnail, logo}`.
- `DELETE /api/projects/<project_key>/image/<variant>` — remove image, clear columns, delete file.
- `GET /project-images/<path:filename>` — serve the stored file with long-lived,
  content-hashed cache headers (`Cache-Control: public, max-age=31536000, immutable`).

`/api/projects` responses gain `thumbnail_url`, `logo_url` (null when unset), plus width/height so
the client can reserve layout space and avoid CLS (cumulative layout shift).

### 5.3 Server-side validation & normalization
- **Accepted input:** PNG, JPEG, WEBP, SVG. (SVG sanitized — strip scripts/external refs — or
  rasterized; see Security.)
- **Max upload size:** 2 MB per file (reject larger with a clear message).
- **Output format (DECIDED): normalize to WEBP by default; fall back to PNG when the source has an
  alpha/transparency channel** (so transparent logos never get a solid box behind them). JPEG/opaque
  sources → WEBP; PNG/SVG/WEBP with transparency → PNG. This keeps files small (WEBP) without ever
  breaking transparency.
- **Normalization (Pillow):**
  - Thumbnail → center-cover crop to square, resize to **256×256**, encode per the format rule above
    (WEBP q82 / PNG).
  - Logo → constrain to max **height 128px**, max **width 512px**, preserve aspect ratio (no padding),
    encode per the format rule above (WEBP q85 / PNG for transparency).
  - Reject decode failures / zero-dimension / absurd aspect ratios (thumbnail must be near-square:
    0.8–1.25; logo must be wider than tall: aspect ≥ 1.6).

- **Auto-thumbnail from logo (DECIDED):** if a project has a **logo but no thumbnail**, the server
  automatically derives a thumbnail by **center-square-cropping the logo** (then applying the 256×256
  normalization above). This runs at logo-upload time when the thumbnail slot is empty, and also on
  read as a safety net. An explicitly uploaded thumbnail always wins and is never overwritten by the
  auto-derived one. Removing the explicit thumbnail re-enables auto-derivation. The stored auto
  thumbnail is flagged (see `thumbnail_is_auto` below) so the UI can show it as "auto from logo".
- Compute `sha256` of the normalized bytes for the filename + cache-busting.
- Write atomically (temp file + rename); delete the previous file for that slot after DB update.

---

## 6. Frontend — Project Settings UI

### 6.1 Where
Each project row/card in `/settings/projects` gains an **"Images"** area with two upload slots side
by side: **Thumbnail (square)** and **Logo (horizontal)**.

### 6.2 UI/UX & beautification (this is a first-class part of the spec)

**Layout & rhythm**
- Two-up **drop zones** in a responsive grid: on desktop they sit side by side; on narrow screens
  they stack. Each drop zone's *frame reflects its target shape* — the thumbnail zone is a perfect
  square, the logo zone is a wide 3:1 rectangle — so the operator intuitively understands the
  expected artwork before reading a single label.
- Consistent **8px spacing scale**; the images area is visually separated from the color/name fields
  with a hairline divider (`1px` at 8% foreground) and a small section caption "Project Imagery".

**The drop zone (empty state)**
- Dashed 1.5px border, `12px` radius, subtle surface tint (`color-mix(in srgb, var(--accent) 6%, transparent)`).
- Centered **Material Symbol** `add_photo_alternate`, a one-line prompt ("Drag a square image or
  click to upload"), and a muted hint line ("PNG, JPG, WEBP or SVG · up to 2 MB · square works
  best"). The hint for the logo zone reads "wide/horizontal works best".
- Hover / drag-over: border brightens to `var(--accent)`, background tint deepens, a soft
  `box-shadow` lift (`0 4px 16px rgba(0,0,0,.12)`) and a **150ms ease** transition. The whole zone
  scales to `1.01` on drag-over to signal "drop target active".
- Keyboard accessible: the zone is a `<button>`/`role="button"` with `tabindex=0`, Enter/Space opens
  the file picker; a visible focus ring (`2px` accent outline, `2px` offset).

**The drop zone (filled state)**
- Shows the actual image inside its correctly-shaped frame with `object-fit: cover` (thumbnail) /
  `object-fit: contain` (logo, so wide art isn't cropped), on a **checkerboard transparency
  backdrop** so transparent PNG/SVG edges read correctly in both themes.
- A translucent **hover toolbar** slides up from the bottom of the frame with two icon buttons:
  `swap_horiz` (Replace) and `delete` (Remove). Icons have `44×44` hit targets and tooltips.
- Below the frame: a tiny metadata line — dimensions + file type + "updated 2h ago" — in
  `12px` muted text.

**Upload feedback**
- On selection: instant **local preview** (object URL) while the request is in flight — no blank
  wait.
- **Progress**: a thin determinate bar across the top edge of the frame (accent color), plus the
  Replace/Remove toolbar disabled during upload.
- **Success**: brief checkmark pulse (`scale` 0.9→1, 200ms) and a toast "Thumbnail updated".
- **Error**: the frame border flashes to `var(--danger)`, an inline message states the exact reason
  (e.g. "Image is 3.1 MB — max is 2 MB", "Logo should be wider than tall"), and the previous image
  is retained (no destructive failure). Errors are announced via `aria-live="assertive"`.

**Live preview strip**
- Under the two zones, a **"Where this appears"** mini-preview renders a mock product tab (using the
  thumbnail) and a mock section header (using the logo) with the project's color as accent — so the
  operator sees the real downstream result before leaving the page. This closes the feedback loop and
  is the single biggest UX win.

**Auto-thumbnail affordance**
- When a logo is uploaded but no explicit thumbnail exists, the thumbnail slot fills with the
  auto-cropped preview and shows a small "Auto from logo" badge (muted pill, top-left of the frame).
- The badge's toolbar offers **Upload your own** (replace) — uploading an explicit thumbnail removes
  the badge; removing that explicit thumbnail returns to the auto state rather than to empty.

**Micro-polish**
- All transitions `150–200ms ease`; respect `prefers-reduced-motion` (disable scale/slide, keep
  color changes).
- Images lazy-load (`loading="lazy"`) and always carry `width`/`height` to prevent layout shift.
- Empty color-only fallback is shown *inside* the preview strip too, so the operator understands what
  "no image" looks like.

### 6.3 Accessibility
- Every image has meaningful `alt` (`"<Project display name> logo"` / `"<name> thumbnail"`).
- Drop zones fully keyboard-operable; focus-visible rings; `aria-describedby` links the hint text.
- Color is never the only signal (icons + text accompany all states).
- Contrast ≥ 4.5:1 for text, ≥ 3:1 for the drop-zone borders, verified in light and dark.

---

## 7. Frontend — RnD Muscle Utilization product-wise view

**Tabs (tight space → thumbnail):**
- Each project tab shows a **20×20 rounded thumbnail** to the left of the project name, replacing (or
  sitting beside) the current colored initials chip. If no thumbnail: keep the initials chip exactly
  as today. The active tab gets the existing accent underline; the thumbnail gets a `1px` ring in the
  project color so it stays legible on any artwork.

**Section header / banner (wide space → logo):**
- When a specific product tab is active, its **logo** renders in the section header at up to `40px`
  height, left-aligned, with the project color as a subtle left border accent. If no logo: fall back
  to the display name in the current header style.

**Consistency:** the same color/initials fallback system already in the codebase remains the source
of truth when images are absent, so nothing regresses for un-imaged projects.

---

## 8. Data flow

1. **Upload:** operator drops a file → client validates type/size, shows local preview → `POST`
   multipart to `/api/projects/<key>/image/<variant>`.
2. **Server:** validate → normalize/re-encode (Pillow) → compute sha256 → write file atomically to
   the persistent image dir → update `managed_projects` metadata columns → delete old file → return
   `{ url, width, height, mime, updated_at }`.
3. **Read (settings):** `/api/projects` returns image URLs + dims → UI renders filled drop zones and
   the live preview strip.
4. **Read (RnD view):** `list_rnd_muscle_project_tabs` / page-state load join `managed_projects` by
   `project_key` → `RndMuscleProjectTab` carries `thumbnail_url`/`logo_url` → tabs render thumbnails,
   header renders logo.
5. **Serve:** `GET /project-images/<file>` streams bytes with immutable cache headers; content-hashed
   filename guarantees correct cache invalidation on replace.

---

## 9. Security & privacy
- **Content sniffing:** validate real image content (Pillow decode), not just extension/MIME header.
- **SVG:** sanitize (remove `<script>`, `on*` handlers, external `href`/`xlink`) or rasterize to WEBP
  on upload; never serve untrusted raw SVG inline.
- **Path traversal:** `project_key` is already constrained to `^[A-Z0-9_-]+$`; generated filenames are
  server-controlled (sha256 + fixed variant) — never trust client filename.
- **Size/DoS:** hard 2 MB cap enforced before full read; reject oversized streams early.
- **Auth:** upload/delete routes sit behind the same settings authorization as other project mutations.
- No image or path is placed in query strings beyond the server-generated static file URL.

---

## 10. Testing
- **Unit (`managed_projects_registry`):** schema init adds columns; set/clear image updates metadata
  and file lifecycle; normalization output dims/mime; aspect-ratio rejection.
- **API (`test_projects_api.py`):** upload happy path returns url+dims; oversize rejected; bad type
  rejected; delete clears columns + removes file; `/api/projects` surfaces urls.
- **RnD view (`test_rnd_muscle_utilization_api.py`):** project tabs carry image urls when set and fall
  back to color/initials when unset.
- **Migration (`test_report_db_migration.py`):** additive columns produce a valid upgrade plan and
  preserve existing rows.

---

## 11. Rollout & migration notes
- Additive, nullable columns → **backward compatible**; old rows simply have empty image metadata and
  render today's color/initials fallback.
- **Production DB upgrade** (operator-driven per protocol): download prod `assignee_hours_capacity.db`,
  run `python db_migration.py --prod <file> --local assignee_hours_capacity.db --plan-only`, review,
  then execute. Because these are additive columns, the plan is a fast column-add — the 4 GB
  download/upload remains the only slow part, and image bytes never enter that round-trip.
- Uploaded files live under the **persistent** `/home`-based dir on Azure and survive deploys; they
  are **not** part of the git deploy package.

---

## 12. Resolved decisions (previously open)
1. **Azure persistent image directory:** `/home/data/project_images/` (the `/home` share survives
   redeploys), via env var `JIRA_PROJECT_IMAGE_DIR`. **Resolved.**
2. **Image format:** normalize to **WEBP**, fall back to **PNG** when the source has transparency.
   **Resolved.**
3. **Missing thumbnail:** auto-derive by **center-square-cropping the logo**; explicit uploads always
   win; `thumbnail_is_auto` flag tracks the state. **Resolved.**

No open questions remain — design is ready for implementation approval.
