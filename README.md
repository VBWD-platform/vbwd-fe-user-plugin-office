# vbwd-fe-user-plugin-office

**VBWD Office** — the user-facing frontend of a self-hosted, privacy-first document suite for the
[VBWD platform](https://github.com/VBWD-platform).

Checkout path: `vbwd-fe-user/plugins/office/` · plugin package: `office`

## Surfaces

| Route | What it is |
|---|---|
| `/dashboard/office` | **VBWD Space** — the file manager: breadcrumb, lazy folder tree, drag-and-drop upload with real progress, context menu (download / rename / move / versions / trash), preview pane, storage-usage bar |
| `/dashboard/office/doc/:id` | **VBWD Docs** — rich-text editor, autosave, version history, edit-lease banner, AI helper sidebar |
| `/dashboard/office/sheet/:id` | **VBWD Spreadsheets** — virtualised grid, formula bar, multi-sheet tabs, CSV/XLSX import & export |
| `/s/:token` | **The public share page** — deliberately *outside* the authenticated layout |

That last one is the point of the product, and the reason it is its own route: a logged-out
stranger opening a share link must see a page with no burger, no user chrome and no store
bootstrap that assumes a session. It must not 500 for want of a token.

## Registration

A named export (the loader falls back to the first named export carrying `.install`; a default
export would not be found):

```ts
export const officePlugin: IPlugin = { /* … */ };
```

`activate()` registers the burger item through `userNavRegistry`, and every route carries
`meta.requiredUserPermission: 'office.use'` — an **access-level** permission, matching the
backend's `@require_user_permission`, not admin RBAC.

## Testing note that matters

Unit tests here must stub the **transport** (`fetch`), not the API-client method. Mocking
`officeApi.listNodes()` and returning a bare array tests the assumption, not the system: the
backend returns `{"items": [...]}`, and a client that forgot to unwrap it once rendered an
*always-empty* file manager while 26 green unit tests looked on. See
`tests/office-api-envelope.spec.ts`.

## Development

```bash
# from vbwd-fe-user/
npx vitest run plugins/office/
npx vue-tsc --noEmit          # the prod build typechecks tests too
E2E_BASE_URL=http://localhost:8080 npx playwright test plugins/office/
```

The E2E suite self-cleans: every run purges the nodes it created, so repeated runs leave the
vault exactly as they found it.

## Companion repos

- [`vbwd-plugin-office`](https://github.com/VBWD-platform/vbwd-plugin-office) — backend
- [`vbwd-fe-admin-plugin-office`](https://github.com/VBWD-platform/vbwd-fe-admin-plugin-office) — admin

## Licence

BSL 1.1 with a Bitcoin-denominated Additional Use Grant. Change Licence: Apache-2.0.
See [`LICENSE`](./LICENSE).
