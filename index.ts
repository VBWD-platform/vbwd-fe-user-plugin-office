import type { IPlugin, IPlatformSDK } from 'vbwd-view-component';
import { userNavRegistry } from '@/plugins/userNavRegistry';
import en from './locales/en.json';
import de from './locales/de.json';

/**
 * VBWD Office — the self-hosted document-vault bundle (S147). This slice
 * ships VBWD Space (the file manager); the Docs/Sheets routes are reserved
 * placeholders (S147-3/-4).
 *
 * Gated on `office.use`, an fe-user ACCESS-LEVEL permission (NOT an admin
 * RBAC permission) — the S146 lesson: player-facing routes must use
 * `requiredUserPermission`, never the admin `@require_permission` gate.
 */
const OFFICE_USE_PERMISSION = 'office.use';

export const officePlugin: IPlugin = {
  name: 'office',
  version: '0.1.0',
  description: 'VBWD Office — a self-hosted document vault (VBWD Space, Docs, Spreadsheets).',
  _active: false,

  install(sdk: IPlatformSDK) {
    sdk.addTranslations('en', en);
    sdk.addTranslations('de', de);

    sdk.addRoute({
      path: '/dashboard/office',
      name: 'office-space',
      component: () => import('./src/views/SpaceHome.vue'),
      meta: { requiresAuth: true, requiredUserPermission: OFFICE_USE_PERMISSION },
    });

    // S147-2 — named shares addressed to the logged-in user.
    sdk.addRoute({
      path: '/dashboard/office/shared-with-me',
      name: 'office-shared-with-me',
      component: () => import('./src/views/OfficeSharedWithMe.vue'),
      meta: { requiresAuth: true, requiredUserPermission: OFFICE_USE_PERMISSION },
    });

    // S147-2 — the public share view. OUTSIDE the authenticated shell on
    // purpose (`noLayout: true`): no burger, no user chrome, no store
    // bootstrap that assumes a session — this is what a logged-out stranger
    // sees, so it must not depend on anything a session would provide.
    sdk.addRoute({
      path: '/s/:token',
      name: 'office-public-share',
      component: () => import('./src/views/OfficePublicShare.vue'),
      props: true,
      meta: { requiresAuth: false, noLayout: true },
    });

    // S147-3 — VBWD Docs: the rich-text editor, edit lease, and AI helper.
    sdk.addRoute({
      path: '/dashboard/office/doc/:id',
      name: 'office-doc',
      component: () => import('./src/views/OfficeDocEditor.vue'),
      props: true,
      meta: { requiresAuth: true, requiredUserPermission: OFFICE_USE_PERMISSION },
    });

    // S147-4 — VBWD Spreadsheets: the virtualised grid, formula bar,
    // multi-sheet tabs, and import/export. Reuses the SAME edit lease as
    // Docs (widened backend resolution, no new lease mechanism).
    sdk.addRoute({
      path: '/dashboard/office/sheet/:id',
      name: 'office-sheet',
      component: () => import('./src/views/OfficeSheetEditor.vue'),
      props: true,
      meta: { requiresAuth: true, requiredUserPermission: OFFICE_USE_PERMISSION },
    });
  },

  activate() {
    this._active = true;
    // The task sketch specified icon: 'folder', but the shared Icon registry
    // (vbwd-fe-core/src/components/ui/icons.ts) has no 'folder' entry — an
    // unknown name silently falls back to a neutral dot. 'document' is the
    // closest existing generic icon for a file-manager nav item.
    userNavRegistry.register({
      pluginName: 'office',
      to: '/dashboard/office',
      labelKey: 'nav.office',
      icon: 'document',
      testId: 'nav-office',
      requiredUserPermission: OFFICE_USE_PERMISSION,
    });
  },

  deactivate() {
    this._active = false;
    userNavRegistry.unregister('office');
  },
};
