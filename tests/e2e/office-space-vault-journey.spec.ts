/**
 * VBWD Office — VBWD Space vault journey (S147-1).
 *
 * Drives the REAL running stack (fe-user :8080, API :5000) through the full
 * user journey the sprint's Definition of Done asks for: nav -> empty state
 * -> folder -> upload -> byte-identical download -> rename -> move with
 * breadcrumb navigation -> usage bar -> trash.
 *
 * Root already holds a seeded "Welcome" folder (plugins/office/populate_db.py)
 * for this demo user, so root itself is never literally empty. The empty
 * state is proven honestly by creating a brand-new folder and opening it —
 * that also covers "create a folder; it appears" in the same, real flow.
 *
 * Serial + a single shared page/login: same rate-limiting reason
 * playwright.config.ts already runs the whole project on one worker (the
 * login endpoint is rate-limited), and later steps depend on the vault state
 * left behind by earlier ones.
 *
 * Run:
 *   cd vbwd-fe-user && E2E_BASE_URL=http://localhost:8080 \
 *     npx playwright test plugins/office/tests/e2e/office-space-vault-journey.spec.ts
 */
import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  loginAsTestUser,
  fetchOfficeJson,
  formatBytesLikeSpaceHome,
  findOfficeNodeIdByName,
  cleanupOfficeNodes,
  officeNodeRow as nodeRow,
  openOfficeContextMenuFor as openContextMenuFor,
  goToOfficeRoot as goToRoot,
} from './support/office-e2e-helpers';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const RUN_ID = Date.now();
const TARGET_FOLDER_NAME = `E2E Target ${RUN_ID}`;
const RENAMED_FILE_NAME = `e2e-renamed-${RUN_ID}.txt`;

const FIXTURE_PATH = path.join(__dirname, 'fixtures', 'office-e2e-fixture.txt');
const FIXTURE_BYTES = fs.readFileSync(FIXTURE_PATH);
// The upload is fed to the input as an in-memory payload with a RUN-SCOPED name
// (setInputFiles accepts {name, mimeType, buffer}, so the payload's filename is
// ours to choose). Uploading under the fixture's own basename instead would
// accumulate identically-named rows at root across runs and make the row locator
// ambiguous on the second run — self-isolating beats cleaning up after the fact.
const UPLOADED_FILE_NAME = `e2e-upload-${RUN_ID}.txt`;

test.describe.serial('VBWD Office — VBWD Space vault journey', () => {
  let vaultPage: Page;
  // Top-level nodes THIS run creates, tracked as they are created so
  // `afterAll` can purge exactly them (tests self-clean; they never wipe the
  // vault). The uploaded file is tracked separately from the folder because
  // it starts at root and is only moved into the folder in test 6b — if an
  // earlier assertion ever fails and aborts the run before that move, both
  // ids are still purged directly.
  let createdFolderId: string | undefined;
  let uploadedFileId: string | undefined;

  test.beforeAll(async ({ browser }) => {
    const context = await browser.newContext();
    vaultPage = await context.newPage();
    await loginAsTestUser(vaultPage);
  });

  test.afterAll(async () => {
    await cleanupOfficeNodes(vaultPage, [uploadedFileId, createdFolderId]);
    await vaultPage.close();
  });

  test('1 — the burger shows VBWD Office and it opens the Space vault', async () => {
    const navItem = vaultPage.locator('[data-testid="nav-office"]');
    await expect(navItem).toBeVisible({ timeout: 10_000 });
    await navItem.click();
    await vaultPage.waitForURL('**/dashboard/office');
    await expect(vaultPage.locator('[data-testid="office-space-home"]')).toBeVisible({
      timeout: 15_000,
    });
    // Root is not empty (the demo seed's "Welcome" folder), which is exactly
    // why the empty state is proven from a fresh folder in the next step.
    await expect(vaultPage.locator('[data-testid="office-node-item"]').first()).toBeVisible({
      timeout: 15_000,
    });
  });

  test('2/3 — creating a folder shows it at root, and opening it shows the empty state', async () => {
    await vaultPage.locator('[data-testid="office-new-folder-btn"]').click();
    await expect(vaultPage.locator('[data-testid="office-new-folder-modal"]')).toBeVisible();
    await vaultPage.locator('[data-testid="office-new-folder-input"]').fill(TARGET_FOLDER_NAME);
    await vaultPage.locator('[data-testid="office-new-folder-confirm"]').click();
    await expect(vaultPage.locator('[data-testid="office-new-folder-modal"]')).toBeHidden();

    const folderRow = nodeRow(vaultPage, TARGET_FOLDER_NAME);
    await expect(folderRow).toBeVisible({ timeout: 10_000 });
    createdFolderId = await findOfficeNodeIdByName(vaultPage, TARGET_FOLDER_NAME);

    await folderRow.click();
    await expect(vaultPage.locator('[data-testid="office-breadcrumb-item"]')).toHaveCount(2);
    await expect(
      vaultPage.locator('[data-testid="office-breadcrumb-item"]').last(),
    ).toHaveText(TARGET_FOLDER_NAME);

    await expect(vaultPage.locator('[data-testid="office-empty-state"]')).toBeVisible();
    await expect(vaultPage.locator('[data-testid="office-empty-upload-cta"]')).toBeVisible();
    await expect(vaultPage.locator('[data-testid="office-node-list"]')).toHaveCount(0);

    await goToRoot(vaultPage);
  });

  test('4 — uploading a file at root shows it with the correct size, and the usage bar grows by exactly its bytes', async () => {
    const usageBefore = await fetchOfficeJson<{ bytes_used: number; bytes_quota: number }>(
      vaultPage,
      '/api/v1/office/usage',
    );

    await vaultPage.locator('[data-testid="office-upload-input"]').setInputFiles({
      name: UPLOADED_FILE_NAME,
      mimeType: 'text/plain',
      buffer: FIXTURE_BYTES,
    });

    const fixtureRow = nodeRow(vaultPage, UPLOADED_FILE_NAME);
    await expect(fixtureRow).toBeVisible({ timeout: 15_000 });
    uploadedFileId = await findOfficeNodeIdByName(vaultPage, UPLOADED_FILE_NAME);

    const expectedSizeLabel = formatBytesLikeSpaceHome(FIXTURE_BYTES.length);
    await expect(fixtureRow.locator('[data-testid="office-node-size"]')).toHaveText(
      expectedSizeLabel,
    );

    const usageAfter = await fetchOfficeJson<{ bytes_used: number; bytes_quota: number }>(
      vaultPage,
      '/api/v1/office/usage',
    );
    expect(usageAfter.bytes_used - usageBefore.bytes_used).toBe(FIXTURE_BYTES.length);

    // The visible usage bar reflects the same number the API reports.
    await expect(vaultPage.locator('[data-testid="office-usage-bar"]')).toContainText(
      formatBytesLikeSpaceHome(usageAfter.bytes_used),
    );
  });

  test('5 — downloading the file returns byte-identical content', async () => {
    await openContextMenuFor(vaultPage, UPLOADED_FILE_NAME);

    const downloadPromise = vaultPage.waitForEvent('download');
    await vaultPage.locator('[data-testid="office-context-download"]').click();
    const download = await downloadPromise;
    const downloadedPath = await download.path();
    expect(downloadedPath).toBeTruthy();
    const downloadedBytes = fs.readFileSync(downloadedPath as string);

    expect(downloadedBytes.equals(FIXTURE_BYTES)).toBe(true);
  });

  test('6a — renaming the file updates its name in the listing', async () => {
    await openContextMenuFor(vaultPage, UPLOADED_FILE_NAME);
    await vaultPage.locator('[data-testid="office-context-rename"]').click();
    await expect(vaultPage.locator('[data-testid="office-rename-modal"]')).toBeVisible();

    const input = vaultPage.locator('[data-testid="office-rename-input"]');
    await input.fill('');
    await input.fill(RENAMED_FILE_NAME);
    await vaultPage.locator('[data-testid="office-rename-confirm"]').click();
    await expect(vaultPage.locator('[data-testid="office-rename-modal"]')).toBeHidden();

    await expect(nodeRow(vaultPage, RENAMED_FILE_NAME)).toBeVisible();
    await expect(nodeRow(vaultPage, UPLOADED_FILE_NAME)).toHaveCount(0);
  });

  test('6b — moving the file into the folder relocates it, and breadcrumb navigation follows it there', async () => {
    await openContextMenuFor(vaultPage, RENAMED_FILE_NAME);
    await vaultPage.locator('[data-testid="office-context-move"]').click();
    await expect(vaultPage.locator('[data-testid="office-move-modal"]')).toBeVisible();

    const moveTree = vaultPage.locator('[data-testid="office-move-tree"]');
    await moveTree.locator('[data-testid="office-tree-toggle"]').first().click();
    const targetLabel = moveTree
      .locator('[data-testid="office-tree-label"]')
      .filter({ hasText: TARGET_FOLDER_NAME });
    await expect(targetLabel).toBeVisible({ timeout: 10_000 });
    await targetLabel.click();

    await expect(
      vaultPage.locator('[data-testid="office-move-target-label"]'),
    ).toHaveText(TARGET_FOLDER_NAME);
    await vaultPage.locator('[data-testid="office-move-confirm"]').click();
    await expect(vaultPage.locator('[data-testid="office-move-modal"]')).toBeHidden();

    // It left the root listing...
    await expect(nodeRow(vaultPage, RENAMED_FILE_NAME)).toHaveCount(0);

    // ...breadcrumb navigation (root -> folder, via clicking the row) finds it there...
    await nodeRow(vaultPage, TARGET_FOLDER_NAME).click();
    await expect(
      vaultPage.locator('[data-testid="office-breadcrumb-item"]').last(),
    ).toHaveText(TARGET_FOLDER_NAME);
    await expect(nodeRow(vaultPage, RENAMED_FILE_NAME)).toBeVisible();

    // ...and the breadcrumb "root" crumb navigates back out again.
    await goToRoot(vaultPage);
    await expect(nodeRow(vaultPage, TARGET_FOLDER_NAME)).toBeVisible();

    // The sidebar folder-tree jump (a second, independent breadcrumb path —
    // `navigateToNode`) also lands on the same folder with its contents.
    const sidebarTree = vaultPage.locator('[data-testid="office-folder-tree"]');
    await sidebarTree.locator('[data-testid="office-tree-toggle"]').first().click();
    await sidebarTree
      .locator('[data-testid="office-tree-label"]')
      .filter({ hasText: TARGET_FOLDER_NAME })
      .click();
    await expect(
      vaultPage.locator('[data-testid="office-breadcrumb-item"]').last(),
    ).toHaveText(TARGET_FOLDER_NAME);
    await expect(nodeRow(vaultPage, RENAMED_FILE_NAME)).toBeVisible();
  });

  test('8 — trashing the file removes it from the listing', async () => {
    await openContextMenuFor(vaultPage, RENAMED_FILE_NAME);
    await vaultPage.locator('[data-testid="office-context-trash"]').click();
    await expect(nodeRow(vaultPage, RENAMED_FILE_NAME)).toHaveCount(0, { timeout: 10_000 });
  });
});
