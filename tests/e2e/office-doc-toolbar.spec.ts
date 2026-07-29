import path from 'path';
import { fileURLToPath } from 'url';
import { test, expect, type Locator, type Page } from '@playwright/test';

// Same ESM shim the sibling specs use — `__dirname` does not exist here.
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Live proof for the S147-3 toolbar bundle, against a real browser and a real
// backend: image persistence across a reload, a real .csv/.xlsx table import,
// and an actually-openable PDF export — none of which a mocked-API unit test
// can establish.
//
// This DOES run in the suite. An earlier version of this header claimed a
// `zzz-` prefix excluded it from CI; no such convention exists in this repo,
// and playwright.config.ts globs every `plugins/*/tests/e2e/**/*.spec.ts`, so
// the file ran and failed. Run it from the HOST, not the app container — the
// container's musl userland cannot launch Playwright's glibc-linked Chromium.
async function clickAtEndOfContent(content: Locator): Promise<void> {
  // Clicks the LAST child of the editor's content — a reliable way to
  // place the cursor at the end before an insert action, regardless of
  // what that last child is (a paragraph, a table, or an image node view).
  await content.locator('> *').last().click();
}

async function waitForSaved(page: Page): Promise<void> {
  await expect
    .poll(
      async () => (await page.locator('[data-testid="office-doc-save-status"]').textContent())?.trim(),
      { timeout: 15000 },
    )
    .toBe('Saved');
}

test('VBWD Docs — formatting, image reload persistence, table import, PDF export', async ({ page }) => {
  test.setTimeout(90000);
  page.on('pageerror', (err) => console.log('PAGE ERROR:', err.message));

  await page.goto('/login');
  await page.fill('[data-testid="email"]', 'test@example.com');
  await page.fill('[data-testid="password"]', 'TestPass123@');
  await page.click('[data-testid="login-button"]');
  await page.waitForURL('**/dashboard');

  const created = await page.evaluate(async () => {
    const token = localStorage.getItem('auth_token') || '';
    const response = await fetch('/api/v1/office/docs', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: `Doc-Verify-${Date.now()}`, parent_id: null }),
    });
    return response.json();
  });
  const docId = created.id;

  await page.goto(`/dashboard/office/doc/${docId}`);
  await expect(page.locator('[data-testid="office-doc-editor"]')).toBeVisible();
  await page.locator('[data-testid="office-doc-toolbar"]').waitFor();

  const content = page.locator('[data-testid="office-doc-content"] .ProseMirror');

  // --- Formatted text, then a heading ---------------------------------------
  await content.click();
  await page.keyboard.type('Hello formatted world');
  await page.keyboard.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A');
  await page.click('[data-testid="office-doc-tb-bold"]');
  await page.click('[data-testid="office-doc-tb-italic"]');
  await page.click('[data-testid="office-doc-tb-underline"]');
  await expect(content.locator('strong')).toBeVisible();
  await expect(content.locator('em')).toBeVisible();
  await expect(content.locator('u')).toBeVisible();

  await page.selectOption('[data-testid="office-doc-tb-style"]', 'heading-2');
  await expect(content.locator('h2')).toHaveText('Hello formatted world');

  // --- Insert image at the end, then verify it STILL renders after reload --
  await clickAtEndOfContent(content);
  await page.setInputFiles(
    '[data-testid="office-doc-tb-image-input"]',
    path.join(__dirname, 'fixtures', 'verify-image.png'),
  );
  await expect(page.locator('[data-testid="office-doc-asset-image"] img')).toBeVisible({ timeout: 15000 });
  await waitForSaved(page);

  await page.reload();
  await page.locator('[data-testid="office-doc-toolbar"]').waitFor();
  const imageAfterReload = page.locator('[data-testid="office-doc-asset-image"] img');
  await expect(imageAfterReload).toBeVisible({ timeout: 15000 });
  await expect(imageAfterReload).toHaveAttribute('src', /^blob:/);
  await expect(content.locator('h2')).toContainText('Hello formatted world');

  // --- Insert a table from a REAL .csv --------------------------------------
  await clickAtEndOfContent(content);
  await page.setInputFiles(
    '[data-testid="office-doc-tb-table-from-file-input"]',
    path.join(__dirname, 'fixtures', 'verify-table.csv'),
  );
  await expect.poll(() => content.locator('table').count(), { timeout: 15000 }).toBeGreaterThan(0);
  await expect(content.locator('table').first()).toContainText('Widget');
  await expect(content.locator('table').first()).toContainText('9.99');
  await expect(content.locator('h2')).toContainText('Hello formatted world');

  // --- Insert a SECOND table from a REAL .xlsx ------------------------------
  await clickAtEndOfContent(content);
  await page.setInputFiles(
    '[data-testid="office-doc-tb-table-from-file-input"]',
    path.join(__dirname, 'fixtures', 'verify-table.xlsx'),
  );
  await expect.poll(() => content.locator('table').count(), { timeout: 15000 }).toBeGreaterThan(1);
  await expect(content.locator('table').nth(1)).toContainText('Ada');
  await expect(content.locator('table').nth(1)).toContainText('Grace');

  // --- Export a PDF, and prove the download is a REAL, openable PDF --------
  await page.click('[data-testid="office-doc-tb-export"]');
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.click('[data-testid="office-doc-tb-export-pdf"]'),
  ]);
  expect(download.suggestedFilename()).toMatch(/\.pdf$/);
  const downloadPath = await download.path();
  const fs = await import('fs');
  const pdfBytes = fs.readFileSync(downloadPath as string);
  expect(pdfBytes.subarray(0, 5).toString('latin1')).toBe('%PDF-');
});
