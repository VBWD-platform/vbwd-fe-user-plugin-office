import { describe, it, expect } from 'vitest';
import { mount } from '@vue/test-utils';
import OfficeVersionsDialog from '../src/components/OfficeVersionsDialog.vue';

const tFallback = (key: string) => key;

// The backend returns versions in ASCENDING order (`version_no.asc()`,
// office_version_repository.py) — the highest version_no, last in the array,
// is the genuinely current one. A previous bug labelled versions[0] (the
// OLDEST) as "Current" instead.
const ascendingVersions = [
  { id: 'v1', version_no: 1, size_bytes: 5, sha256: 'a', created_by_user_id: 'u1', created_at: '2026-07-27T00:00:00Z' },
  { id: 'v2', version_no: 2, size_bytes: 8, sha256: 'b', created_by_user_id: 'u1', created_at: '2026-07-27T01:00:00Z' },
  { id: 'v3', version_no: 3, size_bytes: 12, sha256: 'c', created_by_user_id: 'u1', created_at: '2026-07-27T02:00:00Z' },
];

function mountDialog(versions = ascendingVersions) {
  return mount(OfficeVersionsDialog, {
    props: { versions },
    global: { mocks: { $t: tFallback } },
  });
}

describe('OfficeVersionsDialog', () => {
  it('displays the newest version (highest version_no) FIRST, labelled Current', () => {
    const wrapper = mountDialog();
    const items = wrapper.findAll('[data-testid="office-versions-item"]');

    // The prop arrives in the backend's ASCENDING order (v1, v2, v3), but
    // the dialog re-orders for display: newest on top, oldest at the bottom
    // — matching the E2E contract (office-versions.spec.ts).
    expect(items[0].text()).toContain('v3');
    expect(items[0].find('.office-versions-current').exists()).toBe(true);
    expect(items[2].text()).toContain('v1');
    expect(items[2].find('.office-versions-current').exists()).toBe(false);
  });

  it('offers a Restore button on every version except the current one', () => {
    const wrapper = mountDialog();
    const restoreButtons = wrapper.findAll('[data-testid="office-versions-restore"]');

    expect(restoreButtons).toHaveLength(2);
  });

  it('emits restore with the version_no of the row clicked (the backend contract), not the row id', async () => {
    const wrapper = mountDialog();
    const restoreButtons = wrapper.findAll('[data-testid="office-versions-restore"]');

    // First Restore button belongs to the newest non-current row (v2, since
    // v3 is current and hidden at the top).
    await restoreButtons[0].trigger('click');

    expect(wrapper.emitted('restore')?.[0]).toEqual([2]);
  });
});
