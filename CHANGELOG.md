# Changelog

## [0.1.0] — 2026-07-28

First release — the user-facing frontend of the VBWD Office bundle (sprint S147).

### Added
- **VBWD Space** (`/dashboard/office`): breadcrumb, lazy recursive folder tree, list/grid toggle,
  drag-and-drop upload with real progress (XHR, so progress events actually fire), context menu,
  versions dialog, preview pane, storage-usage bar, and a genuine empty state.
- **Sharing UI**: share dialog with permission choice, an explicit "Anyone with the link, no
  sign-in required" toggle, expiry and password; the token is shown once; a "Shared with me" view.
- **Public share page** (`/s/:token`), registered outside the authenticated layout so a
  logged-out visitor never hits a session-dependent bootstrap.
- **VBWD Docs** (`/dashboard/office/doc/:id`): rich-text editor on a structured JSON model,
  autosave with stale-version conflict handling, version restore, edit-lease banner with
  take-over, and an AI helper sidebar (propose → accept/discard).
- **VBWD Spreadsheets** (`/dashboard/office/sheet/:id`): virtualised grid, formula bar,
  multi-sheet tabs, number/date/error formatting, CSV & XLSX import/export.
- Burger item via `userNavRegistry`; every route gated on the `office.use` access-level permission.
- Playwright E2E suite covering the Space journey, which self-cleans after every run.

### Fixed
- The API client returned the raw `{items: […]}` envelope instead of unwrapping it, so the file
  manager rendered as permanently empty in a real browser while unit tests passed — they mocked
  the client method and bypassed the envelope. A transport-level regression test now pins it.
- `restoreVersion` sent `version_id`; the backend requires `version_no`.
