import { describe, it, expect, beforeEach, vi } from 'vitest';
import { officePlugin } from '../index';
import { userNavRegistry } from '@/plugins/userNavRegistry';

interface AddedRoute {
  path: string;
  name: string;
  meta?: Record<string, unknown>;
}

function makeSdkSpy() {
  return {
    addRoute: vi.fn(),
    addTranslations: vi.fn(),
    addComponent: vi.fn(),
  };
}

const APP_LOCALES = ['en', 'de'];

describe('office plugin registration', () => {
  beforeEach(() => {
    userNavRegistry.unregister('office');
  });

  it('is a named export honouring the IPlugin contract', () => {
    // The loader falls back to the first named export carrying `.install` —
    // a default export would not be found, so this MUST be a named export.
    expect(officePlugin.name).toBe('office');
    expect(typeof officePlugin.install).toBe('function');
    expect(typeof officePlugin.activate).toBe('function');
    expect(typeof officePlugin.deactivate).toBe('function');
  });

  it('registers translations for en and de', () => {
    const sdk = makeSdkSpy();
    officePlugin.install!(sdk as never);
    const registeredLocales = sdk.addTranslations.mock.calls.map((call) => call[0]);
    for (const locale of APP_LOCALES) {
      expect(registeredLocales).toContain(locale);
    }
  });

  it('registers the Space route and both placeholder doc/sheet routes', () => {
    const sdk = makeSdkSpy();
    officePlugin.install!(sdk as never);
    const routes: AddedRoute[] = sdk.addRoute.mock.calls.map((call) => call[0] as AddedRoute);
    const paths = routes.map((route) => route.path);

    expect(paths).toContain('/dashboard/office');
    expect(paths).toContain('/dashboard/office/doc/:id');
    expect(paths).toContain('/dashboard/office/sheet/:id');
  });

  it('registers "Shared with me" (authenticated) and the public /s/:token route', () => {
    const sdk = makeSdkSpy();
    officePlugin.install!(sdk as never);
    const routes: AddedRoute[] = sdk.addRoute.mock.calls.map((call) => call[0] as AddedRoute);
    const byPath = Object.fromEntries(routes.map((route) => [route.path, route]));

    expect(byPath['/dashboard/office/shared-with-me']).toMatchObject({
      meta: { requiresAuth: true, requiredUserPermission: 'office.use' },
    });
    // S147-2: outside the authenticated layout — a logged-out stranger opens
    // this, so it must not require auth and must not render UserLayout.
    expect(byPath['/s/:token']).toMatchObject({
      meta: { requiresAuth: false, noLayout: true },
    });
  });

  it('gates every AUTHENTICATED office route on auth + the office.use permission', () => {
    const sdk = makeSdkSpy();
    officePlugin.install!(sdk as never);
    const routes: AddedRoute[] = sdk.addRoute.mock.calls.map((call) => call[0] as AddedRoute);
    // The public share view is the one deliberate exception (S147-2 D6-style
    // frontend counterpart) — every other route stays behind the same bar.
    const authenticatedRoutes = routes.filter((route) => route.path !== '/s/:token');

    for (const route of authenticatedRoutes) {
      expect(route.meta?.requiresAuth).toBe(true);
      expect(route.meta?.requiredUserPermission).toBe('office.use');
    }
  });

  it('registers the burger nav item pointing at /dashboard/office', () => {
    officePlugin.activate!();

    const items = Array.from(
      (userNavRegistry as unknown as { getSidebarItems(): Array<Record<string, unknown>> }).getSidebarItems(),
    ).filter((item) => item.pluginName === 'office');

    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({
      pluginName: 'office',
      to: '/dashboard/office',
      labelKey: 'nav.office',
      testId: 'nav-office',
      requiredUserPermission: 'office.use',
    });
  });

  it('deactivate removes the nav item', () => {
    officePlugin.activate!();
    officePlugin.deactivate!();
    const items = userNavRegistry.getSidebarItems().filter((item) => item.pluginName === 'office');
    expect(items).toHaveLength(0);
  });
});
