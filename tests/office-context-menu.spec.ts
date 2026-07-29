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

function mountMenu(x: number, y: number) {
  return mount(OfficeContextMenu, {
    props: { node: NODE, x, y },
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
    const NOMINAL_MENU_HEIGHT = 200;
    const { top } = styleOf(mountMenu(400, 723));

    expect(top + NOMINAL_MENU_HEIGHT).toBeLessThanOrEqual(VIEWPORT_HEIGHT);
    expect(top).toBeGreaterThanOrEqual(0);
  });

  it('never renders past the right edge', () => {
    // The whole menu box must fit, not just its origin: a left of 1275 is
    // "inside" the viewport and still puts every button off-screen.
    const NOMINAL_MENU_WIDTH = 170;
    const { left } = styleOf(mountMenu(1275, 300));

    expect(left + NOMINAL_MENU_WIDTH).toBeLessThanOrEqual(VIEWPORT_WIDTH);
    expect(left).toBeGreaterThanOrEqual(0);
  });

  it('stays on screen even when the pointer is in the very corner', () => {
    const { top, left } = styleOf(mountMenu(VIEWPORT_WIDTH, VIEWPORT_HEIGHT));

    const NOMINAL_MENU_WIDTH = 170;
    const NOMINAL_MENU_HEIGHT = 200;
    expect(top).toBeGreaterThanOrEqual(0);
    expect(top + NOMINAL_MENU_HEIGHT).toBeLessThanOrEqual(VIEWPORT_HEIGHT);
    expect(left).toBeGreaterThanOrEqual(0);
    expect(left + NOMINAL_MENU_WIDTH).toBeLessThanOrEqual(VIEWPORT_WIDTH);
  });

  it('still shows every action for a document', () => {
    const wrapper = mountMenu(400, 300);

    for (const action of ['download', 'share', 'rename', 'move', 'versions', 'trash']) {
      expect(wrapper.find(`[data-testid="office-context-${action}"]`).exists()).toBe(true);
    }
  });
});
