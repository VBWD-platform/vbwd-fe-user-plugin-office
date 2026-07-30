import { describe, it, expect, beforeEach } from 'vitest';
import { mount } from '@vue/test-utils';
import OfficeContextMenu from '../src/components/OfficeContextMenu.vue';
import type { OfficeNode } from '../src/api/officeApi';

/**
 * The menu is `position: fixed` at the pointer, so on a row near the bottom or
 * right edge it can render OUTSIDE the viewport — visible in the DOM, correct
 * in every snapshot, and completely unclickable.
 *
 * That is not hypothetical: it took the versions E2E down. The menu opened at
 * y=723 in a 720px-tall viewport, so `document.elementFromPoint` at the
 * button's centre returned null and the click timed out. Nothing threw, no
 * console error appeared — the feature was simply unreachable for any file low
 * enough in the list.
 */
const NODE: OfficeNode = {
  id: 'n-1',
  name: 'report.txt',
  kind: 'document',
  parent_id: null,
  created_at: null,
  updated_at: null,
} as unknown as OfficeNode;

const VIEWPORT_WIDTH = 1280;
const VIEWPORT_HEIGHT = 720;
const NOMINAL_MENU_WIDTH = 170;
// Mirrors the component: item height x rendered items + vertical padding. A
// DOCUMENT menu renders 9 items, which is what made a fixed 200px estimate
// under-clamp once copy/cut/delete were added.
const DOCUMENT_MENU_ITEMS = 9;
const NOMINAL_MENU_HEIGHT = DOCUMENT_MENU_ITEMS * 34 + 16;

function mountMenu(x: number, y: number, overrides: Record<string, unknown> = {}) {
  return mount(OfficeContextMenu, {
    props: { node: NODE, x, y, canPaste: false, ...overrides },
    global: { mocks: { $t: (key: string) => key } },
  });
}

function styleOf(wrapper: ReturnType<typeof mountMenu>) {
  const style = wrapper.get('[data-testid="office-context-menu"]').attributes('style') || '';
  const top = Number(/top:\s*([\d.-]+)px/.exec(style)?.[1] ?? NaN);
  const left = Number(/left:\s*([\d.-]+)px/.exec(style)?.[1] ?? NaN);
  return { top, left, style };
}

describe('OfficeContextMenu — stays inside the viewport', () => {
  beforeEach(() => {
    Object.defineProperty(window, 'innerWidth', { value: VIEWPORT_WIDTH, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: VIEWPORT_HEIGHT, configurable: true });
  });

  it('renders at the pointer when there is room', () => {
    const { top, left } = styleOf(mountMenu(400, 300));

    expect(top).toBe(300);
    expect(left).toBe(400);
  });

  it('never renders below the bottom edge', () => {
    // The exact failure that broke the versions E2E.
    const { top } = styleOf(mountMenu(400, 723));

    expect(top + NOMINAL_MENU_HEIGHT).toBeLessThanOrEqual(VIEWPORT_HEIGHT);
    expect(top).toBeGreaterThanOrEqual(0);
  });

  it('never renders past the right edge', () => {
    // The whole menu box must fit, not just its origin: a left of 1275 is
    // "inside" the viewport and still puts every button off-screen.
    const { left } = styleOf(mountMenu(1275, 300));

    expect(left + NOMINAL_MENU_WIDTH).toBeLessThanOrEqual(VIEWPORT_WIDTH);
    expect(left).toBeGreaterThanOrEqual(0);
  });

  it('stays on screen even when the pointer is in the very corner', () => {
    const { top, left } = styleOf(mountMenu(VIEWPORT_WIDTH, VIEWPORT_HEIGHT));

    expect(top).toBeGreaterThanOrEqual(0);
    expect(top + NOMINAL_MENU_HEIGHT).toBeLessThanOrEqual(VIEWPORT_HEIGHT);
    expect(left).toBeGreaterThanOrEqual(0);
    expect(left + NOMINAL_MENU_WIDTH).toBeLessThanOrEqual(VIEWPORT_WIDTH);
  });

  it('still shows every action for a document', () => {
    const wrapper = mountMenu(400, 300);

    for (const action of ['download', 'share', 'rename', 'copy', 'cut', 'move', 'versions', 'trash', 'delete']) {
      expect(wrapper.find(`[data-testid="office-context-${action}"]`).exists()).toBe(true);
    }
  });

  it('a node-context menu never offers Paste — that lives in the empty-space menu', () => {
    const wrapper = mountMenu(400, 300);

    expect(wrapper.find('[data-testid="office-context-paste"]').exists()).toBe(false);
  });

  it('emits the action name for copy/cut/delete like every other item', async () => {
    const wrapper = mountMenu(400, 300);

    await wrapper.find('[data-testid="office-context-copy"]').trigger('click');
    expect(wrapper.emitted('action')?.[0]).toEqual(['copy']);

    await wrapper.find('[data-testid="office-context-cut"]').trigger('click');
    expect(wrapper.emitted('action')?.[1]).toEqual(['cut']);

    await wrapper.find('[data-testid="office-context-delete"]').trigger('click');
    expect(wrapper.emitted('action')?.[2]).toEqual(['delete']);
  });

  it('empty-space (no node) offers Paste only, disabled without a clipboard', () => {
    const wrapper = mountMenu(400, 300, { node: null, canPaste: false });

    expect(wrapper.find('[data-testid="office-context-paste"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-context-paste"]').attributes('disabled')).toBeDefined();
    for (const action of ['download', 'share', 'rename', 'copy', 'cut', 'move', 'versions', 'trash', 'delete']) {
      expect(wrapper.find(`[data-testid="office-context-${action}"]`).exists()).toBe(false);
    }
  });

  it('empty-space Paste is enabled once the clipboard has an entry', async () => {
    const wrapper = mountMenu(400, 300, { node: null, canPaste: true });

    expect(wrapper.find('[data-testid="office-context-paste"]').attributes('disabled')).toBeUndefined();
    await wrapper.find('[data-testid="office-context-paste"]').trigger('click');
    expect(wrapper.emitted('action')?.[0]).toEqual(['paste']);
  });
});
