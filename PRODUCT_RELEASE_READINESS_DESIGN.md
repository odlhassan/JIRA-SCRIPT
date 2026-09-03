# Product Release Readiness Design

## Purpose

The release readiness page turns the existing Product Releases planner into a quick
collaborative board. Products, releases, release fields, epics, and checklist work are
kept in one focused path so the board stays as fast to update as the PowerPoint source.

Release number, release date, epic assignments, and completion use the existing Product
Releases APIs. Checklist readiness, notes, and the prototype archive are browser-local
until a persistent checklist model is approved.

## Business Logic

- Products are vertical tabs at the left of the page.
- Product identity is authoritative to active managed projects and live release records.
  A demonstration release is included only when its project key is recognized by one of
  those sources, so sample data cannot create a product the user did not configure.
- Selecting a product shows its active releases; selecting a release opens its details.
- Release number and release date are editable in place. Live records use the same
  Product Releases update endpoint and validation as `/settings/product-releases`.
- Selecting Done for a live release opens an inline completion form. Actual release date
  and completed-by actor are required, matching Product Releases. Confirming records a
  `released` action, actual date, actor, optional comments, lifecycle status, and history
  through the shared release-action endpoint.
- A release marked Released in Product Releases appears as Done on this board. Moving a
  released item to another readiness status records the shared `reverted` action first,
  keeping the release lifecycle and readiness completion state consistent.
- Epics are never typed manually. The Add epics dialog searches and multi-selects from
  the existing epic pool. Removing an epic also uses the existing release assignment API.
- Release responsible is intentionally absent. Responsible remains available for epics,
  checklists, and checklist content.
- The same four-state control applies to releases, epics, checklists, and content:
  Done, Planned, Skipped, and Need confirmation.
- The current status is yellow. Planned can additionally carry a manual Delayed flag,
  which changes only Planned to red. Selecting another status clears Delayed.
- Need confirmation requires two names: who takes confirmation and from whom it is needed.
- Checklist and content titles are edited directly on their visible cards.
- Checklist content scope is either Whole release or one or more assigned epics. The
  scope menu supports search and multi-selection using epic key and name.
- No date is tracked on a checklist or checklist-content row.
- Links, evidence, and notes remain optional secondary actions in the details drawer.
- Archive appears only when the selected release is Done. For live releases, Done is
  available only after the Product Releases completion action succeeds. Archiving removes
  the release from the active list and places it in the bottom Archived releases area,
  where it can be restored. Archive placement remains browser-local.
- If the API is unavailable, demonstration releases remain usable for design review.

## Business Cases

- A release manager switches products on the left and sees only that product's releases.
- OmniConnect appears under its configured `O2` / `OmniConnect-2025` project identity;
  the legacy demonstration key `OMNICONNECT` does not create a second product tab.
- During a release meeting, the release number or date is corrected without leaving the board.
- A team searches the database epic pool and assigns several epics in one action.
- Documentation applies to the whole release while individual feature-video rows target
  several named epics.
- Customer Success owns Stakeholder buy-in while Hassan is recorded as the person taking
  confirmation from an implementation or customer team.
- A Planned announcement is manually marked delayed without inventing a checklist due date.
- A completed release is archived so current work remains prominent, then restored if needed.
- A release completed from either Product Releases or this board shows the same Released
  lifecycle state and action history.

## Interaction Model

The main screen has three working areas: product tabs, the selected product's release
list, and the selected release's details. Release editing, epic assignment, status,
responsible person, confirmation, scope, title edits, and add/remove actions are all on
the visible surface. Detail and checklist rows use consistent responsive grid columns for
identity, status, owner, scope, and actions; narrow layouts stack the same columns without
changing their order. The details drawer is reserved for optional notes and evidence.

Status menus display words and icons rather than relying on color alone. Scope selection
uses checkboxes so a content row can cover multiple epics, while Whole release remains a
mutually exclusive option.

## Front-end UI Fields

- **Product tab:** filters the active and archived releases by product.
- **Release item:** selects the release shown in the details area.
- **Release number:** updates the selected live release through the Product Releases API.
- **Release date:** uses the existing API's validation and update behavior.
- **Save release:** persists both release fields for live records; demonstration records
  update only in the prototype.
- **Release status:** changes browser-local readiness status.
- **Complete this release:** appears after choosing Done on a live planned release.
- **Actual release date:** required date passed to the shared Released action.
- **Completed by:** required actor recorded in Product Releases history.
- **Comments:** optional completion note recorded with the Released action.
- **Lifecycle sync note:** distinguishes local design data from live completion and shows
  when a release is synchronized as Released.
- **Archive / Restore:** moves a Done release between active and archived prototype lists.
- **Add epics from database:** opens a searchable, multi-select epic pool dialog.
- **Remove epic:** removes an existing release-to-epic assignment.
- **Responsible:** selects the person or team for an epic, checklist, or content row.
- **Done / Planned / Skipped / Need confirmation:** shared status choices at all levels.
- **Mark delayed:** manual flag available only for Planned.
- **Confirmation taken by / from:** required pair shown for Need confirmation.
- **Scope:** searchable multi-select of Whole release or assigned epic names.
- **Editable title:** direct text editing for checklists and content.
- **Add checklist / Add content / Remove:** direct structure changes on the board.
- **Ellipsis:** opens optional link/evidence and notes.

## Script Files

- `product_release_readiness_design.html` defines the responsive product rail, release
  list, aligned detail/checklist grids, archive, epic picker, status and scope popovers,
  and details drawer.
- `product-release-readiness-design.js` loads live release and epic data, updates live
  release fields, assignments, and lifecycle actions; synchronizes Released with Done;
  manages the shared readiness model; and persists checklist and archive state locally.
- `report_server.py` registers the route, serves its JavaScript asset, and links the
  design from Product Releases.
- `tests/test_product_releases_api.py` verifies the route, asset, entry link, navigator,
  database epic controls, status/confirmation/scope markers, archive, and absence of dates.

## Dependent & Impacted Files

- `README.md` documents the localhost route and live-versus-local behavior.
- `EXPECTED_FILES.md` lists the HTML and JavaScript required by the route.
- `report_html/shared-nav.css`, `report_html/shared-nav.js`, and
  `report_html/material-symbols.css` provide the implemented project's navigation and icons.
- Existing Product Releases endpoints and tables are reused without a schema change.

## Table Schema

No SQLite schema change is included.

The page reads current data backed by `product_releases`, `product_release_epics`, and
`epics_management`. The proposed persistent readiness entity is still one reusable shape
for release, epic, checklist, and content status, with optional scope, ownership,
confirmation, evidence, and notes. Persistence remains deferred until the interaction
design is approved.

## Data Flow

1. The readiness route serves the HTML and JavaScript.
2. JavaScript loads releases, the epic pool, and project display names.
3. Live release keys and active managed-project keys form the recognized product set.
   Demonstration releases outside that set are discarded before rendering tabs.
4. Selecting a product filters releases; selecting a release renders its board.
5. Saving live release number/date sends `PUT /api/product-releases/<release_id>`.
6. Adding epics sends `POST /api/product-releases/<release_id>/epics`; removing sends
   `DELETE /api/product-releases/<release_id>/epics/<epic_row_id>`.
7. Completing a live release sends `POST /api/product-releases/<release_id>/actions` with
   action `released`, actual date, actor, and optional notes. The returned lifecycle and
   action history become the visible Done state.
8. Loading a live release reconciles an externally changed Released lifecycle into Done.
9. Readiness, checklist structure, notes, and archive state update the browser-local model.
10. All affected visible controls rerender from that shared model.

## Change Notes

- 2026-09-03: Made managed/live project keys authoritative for product tabs so the
  unconfigured `OMNICONNECT` demo key cannot duplicate `O2` / `OmniConnect-2025`.
- 2026-09-03: Aligned release details, epics, checklist headers, and checklist content on
  consistent responsive grids. Connected Done and reopen behavior to the same release
  action/history API used by Product Releases.
- 2026-09-03: Replaced the multi-view concept with a focused product → release → details
  navigator. Added inline live release editing, database-backed epic add/remove, searchable
  multi-epic scope, simplified checklist cards, and a Done-only browser-local archive.
- 2026-09-03: Removed checklist dates; added shared inline status, manual Planned delay,
  ownership, two-party confirmation, direct editing, flexible scope, and optional details.
- 2026-09-03: Added the initial release readiness design route.
