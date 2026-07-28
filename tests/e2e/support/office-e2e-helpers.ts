/**
 * Shared helpers for the VBWD Office (S147-1) E2E suite.
 *
 * Kept local to `plugins/office/tests/e2e/` (SOLID — the plugin owns its own
 * E2E coverage; it does not reach into host-app test infrastructure), but
 * shared BETWEEN this plugin's own spec files so the login flow, the
 * auth-token-aware fetch helper, and the byte-formatting rule used to build
 * assertion strings each have exactly one home (DRY).
 */
import { expect, type Page } from '@playwright/test';

export const TEST_USER = { email: 'test@example.com', password: 'TestPass123@' };

export async function loginAsTestUser(page: Page): Promise<void> {
  await page.goto('/login');
  await page.fill('[data-testid="email"]', TEST_USER.email);
  await page.fill('[data-testid="password"]', TEST_USER.password);
  await page.click('[data-testid="login-button"]');
  await page.waitForURL('**/dashboard');
}

/**
 * Call an `/api/v1/office/*` endpoint from inside the page's own origin,
 * carrying the same `auth_token` the app's fetch calls use (the officeApi
 * client reads it from localStorage — see
 * `vbwd-fe-user/plugins/office/src/api/officeApi.ts`). This exercises the
 * REAL backend over the REAL browser fetch stack; it is not a mocked call.
 */
export async function fetchOfficeJson<T>(
  page: Page,
  path: string,
  init?: { method?: string; body?: unknown },
): Promise<T> {
  return page.evaluate(
    async ({ requestPath, requestInit }) => {
      const token = localStorage.getItem('auth_token') || '';
      const response = await fetch(requestPath, {
        method: requestInit?.method ?? 'GET',
        headers: {
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
          ...(requestInit?.body !== undefined ? { 'Content-Type': 'application/json' } : {}),
        },
        body: requestInit?.body !== undefined ? JSON.stringify(requestInit.body) : undefined,
      });
      return response.json();
    },
    { requestPath: path, requestInit: init },
  );
}

/**
 * Upload a new document VERSION straight through the real
 * `POST /documents/<id>/versions` endpoint, built and sent entirely inside
 * the browser (a real multipart HTTP request, not a mock).
 *
 * This exists because — as of this E2E pass — `SpaceHome.vue` has NO user
 * action wired to that endpoint (see the walkthrough report): dropping a
 * file with the same name onto an existing document creates a SIBLING
 * document, not a new version of it. Exercising the endpoint this way keeps
 * the versions-dialog test honest against real backend history without
 * inventing a UI path that does not exist.
 */
export async function uploadOfficeVersionViaApi(
  page: Page,
  documentId: string,
  fileName: string,
  content: string,
): Promise<{ status: number; body: unknown }> {
  return page.evaluate(
    async ({ id, name, text }) => {
      const token = localStorage.getItem('auth_token') || '';
      const formData = new FormData();
      formData.append('file', new Blob([text], { type: 'text/plain' }), name);
      const response = await fetch(`/api/v1/office/documents/${id}/versions`, {
        method: 'POST',
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        body: formData,
      });
      const body = await response.json().catch(() => null);
      return { status: response.status, body };
    },
    { id: documentId, name: fileName, text: content },
  );
}

/**
 * Locate a node row by its EXACT displayed name (never a substring match —
 * run-scoped fixture names deliberately share long common prefixes across
 * this suite, e.g. `office-e2e-fixture.txt` vs `office-e2e-fixture-v2.txt`).
 */
export interface OfficeNodeListItem {
  id: string;
  name: string;
  kind: 'folder' | 'document';
}

/**
 * Resolve a top-level (root) node's id by its exact name, through the real
 * `GET /api/v1/office/nodes` listing. Used both by tests that need the id of
 * something they just created through the UI, and by this file's own cleanup
 * helper below.
 */
export async function findOfficeNodeIdByName(page: Page, exactName: string): Promise<string> {
  const { items } = await fetchOfficeJson<{ items: OfficeNodeListItem[] }>(
    page,
    '/api/v1/office/nodes',
  );
  const match = items.find((item) => item.name === exactName);
  if (!match) throw new Error(`office node "${exactName}" not found in root listing`);
  return match.id;
}

/**
 * Permanently delete ONE node this run created, through the real,
 * authenticated `DELETE /api/v1/office/nodes/<id>?purge=true` endpoint
 * (`purge=true` — a plain trash would leave the row behind). A 404 is treated
 * as success: the node is already gone, which is the desired end state.
 */
async function purgeOfficeNode(page: Page, nodeId: string): Promise<void> {
  const status = await page.evaluate(async (id) => {
    const token = localStorage.getItem('auth_token') || '';
    const response = await fetch(`/api/v1/office/nodes/${id}?purge=true`, {
      method: 'DELETE',
      headers: token ? { Authorization: `Bearer ${token}` } : {},
    });
    return response.status;
  }, nodeId);
  if (status !== 204 && status !== 404) {
    throw new Error(`DELETE /api/v1/office/nodes/${nodeId}?purge=true -> ${status}`);
  }
}

/**
 * Best-effort cleanup for the TOP-LEVEL nodes a run created directly.
 *
 * Rule on this platform: tests self-clean, they never wipe the database —
 * this deletes only the ids the caller passes in (nodes THIS run created),
 * never a broad sweep. Purging a folder cascades to everything inside it
 * (including already-trashed children — see
 * `DocumentService._purge_recursive`), so a caller that tracks the top-level
 * folder/document ids it created does not also need to track what later
 * moved inside them.
 *
 * A failure to clean up one node is logged and skipped, never thrown — this
 * must never fail an otherwise-passing test run.
 */
export async function cleanupOfficeNodes(
  page: Page,
  nodeIds: Array<string | null | undefined>,
): Promise<void> {
  for (const nodeId of nodeIds) {
    if (!nodeId) continue;
    try {
      await purgeOfficeNode(page, nodeId);
    } catch (error) {
      // Cleanup is best-effort by design (see doc comment above) — log
      // clearly so litter is still visible, but never fail the run over it.
      console.warn(`[office-e2e cleanup] failed to purge node ${nodeId}:`, error);
    }
  }
}

export function officeNodeRow(page: Page, exactName: string) {
  return page.locator('[data-testid="office-node-item"]').filter({
    has: page.getByTestId('office-node-name').getByText(exactName, { exact: true }),
  });
}

export async function openOfficeContextMenuFor(page: Page, exactName: string): Promise<void> {
  await officeNodeRow(page, exactName).locator('[data-testid="office-node-menu-btn"]').click();
  await expect(page.locator('[data-testid="office-context-menu"]')).toBeVisible();
}

export async function goToOfficeRoot(page: Page): Promise<void> {
  await page.locator('[data-testid="office-breadcrumb-item"]').first().click();
  await expect(page.locator('[data-testid="office-breadcrumb-item"]')).toHaveCount(1);
}

/** Mirrors `SpaceHome.vue`'s own `formatBytes` so assertions can build the
 * exact label text the UI is expected to render, without hard-coding a
 * magic byte count in every test. */
export function formatBytesLikeSpaceHome(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(1)} ${units[unitIndex]}`;
}
