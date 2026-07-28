/**
 * VBWD Office — versions dialog + restore (S147-1 acceptance #7).
 *
 * Kept in its own file (rather than the main journey spec) because it
 * exercises a genuine product gap: `SpaceHome.vue` has no user action wired
 * to `POST /documents/<id>/versions` (uploading over an existing document
 * creates a SIBLING document, not a new version), so the second version is
 * created with a direct, real, authenticated fetch from inside the page —
 * see `uploadOfficeVersionViaApi` for exactly why. The dialog itself, and the
 * restore action's WIRE CONTRACT, are still exercised through the real UI.
 *
 * Run:
 *   cd vbwd-fe-user && E2E_BASE_URL=http://localhost:8080 \
 *     npx playwright test plugins/office/tests/e2e/office-versions.spec.ts
 */
import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  loginAsTestUser,
  uploadOfficeVersionViaApi,
  findOfficeNodeIdByName,
  cleanupOfficeNodes,
  officeNodeRow as nodeRow,
  openOfficeContextMenuFor as openContextMenuFor,
} from './support/office-e2e-helpers';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const RUN_ID = Date.now();
const DOCUMENT_NAME = `e2e-versions-${RUN_ID}.txt`;
const VERSION_2_FIXTURE_PATH = path.join(__dirname, 'fixtures', 'office-e2e-fixture-v2.txt');
const VERSION_2_CONTENT = fs.readFileSync(VERSION_2_FIXTURE_PATH, 'utf8');

// `.serial`, deliberately: these tests share ONE manually-created `page`
// across the whole file (not the per-test `page` fixture), and a test that
// times out while it shares that page with a NON-serial sibling can tear the
// browser context down out from under the next test (Playwright's leaked-page
// cleanup on a worker-level timeout), which then reports a confusing
// "context has been closed" instead of a clean, readable SKIP. `.serial`
// avoids that: once one test in this chain fails for real, the rest report
// as skipped rather than as a second, misleading failure.
test.describe.serial('VBWD Office — versions dialog and restore', () => {
  let versionsPage: Page;
  let documentId: string;

  test.beforeAll(async ({ browser }) => {
    const context = await browser.newContext();
    versionsPage = await context.newPage();
    await loginAsTestUser(versionsPage);
    // Client-side navigation (not `page.goto`, which does a full reload):
    // fe-user plugins install asynchronously after the shell mounts, and a
    // hard reload straight onto a plugin route can race that and hit the
    // router's catch-all 404 before `office`'s route is registered.
    await versionsPage.locator('[data-testid="nav-office"]').click();
    await versionsPage.waitForURL('**/dashboard/office');
    await expect(versionsPage.locator('[data-testid="office-space-home"]')).toBeVisible({
      timeout: 15_000,
    });
  });

  test.afterAll(async () => {
    // Tests self-clean; they never wipe the vault — purge only the ONE
    // top-level document this run created (its version history purges with
    // it, see `DocumentService._purge_recursive`).
    await cleanupOfficeNodes(versionsPage, [documentId]);
    await versionsPage.close();
  });

  test('setup — upload a document, then a real second version via the API', async () => {
    // Version 1, through the real upload UI.
    const dataTransferPath = path.join(__dirname, 'fixtures', 'office-e2e-fixture.txt');
    const fileChooserInput = versionsPage.locator('[data-testid="office-upload-input"]');
    await fileChooserInput.setInputFiles({
      name: DOCUMENT_NAME,
      mimeType: 'text/plain',
      buffer: fs.readFileSync(dataTransferPath),
    });
    await expect(nodeRow(versionsPage, DOCUMENT_NAME)).toBeVisible({ timeout: 15_000 });

    documentId = await findOfficeNodeIdByName(versionsPage, DOCUMENT_NAME);

    // Version 2, through the real backend endpoint (see file header — no UI
    // action reaches it today).
    const result = await uploadOfficeVersionViaApi(
      versionsPage,
      documentId,
      DOCUMENT_NAME,
      VERSION_2_CONTENT,
    );
    expect(result.status).toBe(201);
  });

  test('the versions dialog lists real history: v2 current, v1 restorable', async () => {
    await openContextMenuFor(versionsPage, DOCUMENT_NAME);
    await versionsPage.locator('[data-testid="office-context-versions"]').click();
    await expect(versionsPage.locator('[data-testid="office-versions-modal"]')).toBeVisible();

    const items = versionsPage.locator('[data-testid="office-versions-item"]');
    await expect(items).toHaveCount(2, { timeout: 10_000 });

    // Correct, un-weakened expectation: the newest version (v2, just
    // uploaded) is "current", and the older v1 is the one offered for
    // restore.
    //
    // This currently fails: `GET /documents/<id>/versions`
    // (OfficeVersionRepository.find_by_document_id) returns versions in
    // ASCENDING version_no order (v1, v2, …), but
    // `OfficeVersionsDialog.vue` labels `versions[0]` "Current" and puts the
    // Restore button on every OTHER row — i.e. it assumes descending order.
    // The result: v1 (the oldest) is shown as "Current" and v2 (the actual
    // current version) gets a "Restore" button. Reported as a real product
    // defect — not weakened here.
    await expect(items.first()).toContainText('v2');
    await expect(items.first().locator('.office-versions-current')).toBeVisible();
    await expect(items.last()).toContainText('v1');
    await expect(
      items.last().locator('[data-testid="office-versions-restore"]'),
    ).toBeVisible();
  });

  test('restoring v1 appends a new current version', async () => {
    const items = versionsPage.locator('[data-testid="office-versions-item"]');
    const responsePromise = versionsPage.waitForResponse(
      (response) =>
        response.url().includes(`/documents/${documentId}/restore`) &&
        response.request().method() === 'POST',
    );
    await items.last().locator('[data-testid="office-versions-restore"]').click();
    const response = await responsePromise;

    // Correct, un-weakened expectation per S147-1 ("restore appends rather
    // than mutating; history stays complete"): a successful restore returns
    // 201 and the dialog grows to 3 entries with the restored copy on top.
    //
    // This currently fails: `officeApi.restoreVersion` (src/api/officeApi.ts)
    // POSTs `{ version_id: <uuid> }`, but the backend route
    // (plugins/office/office/routes.py::restore_document_version) requires
    // an integer `{ version_no }`. The request 400s and the dialog never
    // updates. Reported as a real product defect — not weakened here.
    expect(response.status()).toBe(201);
    await expect(items).toHaveCount(3, { timeout: 10_000 });
  });
});
