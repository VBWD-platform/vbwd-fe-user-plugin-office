import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises, RouterLinkStub } from '@vue/test-utils';

const mockListSharedWithMe = vi.fn();
const mockDownloadSharedDocument = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    listSharedWithMe: (...args: unknown[]) => mockListSharedWithMe(...args),
    downloadSharedDocument: (...args: unknown[]) => mockDownloadSharedDocument(...args),
  },
}));

vi.mock('vbwd-view-component', () => ({
  Icon: { name: 'IconStub', props: ['name', 'size'], template: '<span :data-icon="name" />' },
}));

import OfficeSharedWithMe from '../src/views/OfficeSharedWithMe.vue';

const tFallback = (key: string) => key;

function mountView() {
  return mount(OfficeSharedWithMe, {
    global: {
      mocks: { $t: tFallback },
      stubs: { RouterLink: RouterLinkStub },
    },
  });
}

const sharedEntry = {
  id: 'd1',
  parent_id: null,
  kind: 'document' as const,
  name: 'Roadmap.pdf',
  trashed_at: null,
  created_at: '2026-07-27T00:00:00Z',
  updated_at: '2026-07-27T00:00:00Z',
  share: {
    id: 's1',
    node_id: 'd1',
    created_by_user_id: 'owner-1',
    permission: 'view' as const,
    subject_user_id: 'me',
    allow_anonymous: false,
    has_password: false,
    expires_at: null,
    revoked_at: null,
    last_used_at: null,
    created_at: '2026-07-27T00:00:00Z',
    updated_at: '2026-07-27T00:00:00Z',
  },
};

describe('OfficeSharedWithMe', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('renders the required data-testid on the root container', async () => {
    mockListSharedWithMe.mockResolvedValue([]);
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-shared-with-me"]').exists()).toBe(true);
  });

  it('shows an empty state when nothing has been shared', async () => {
    mockListSharedWithMe.mockResolvedValue([]);
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-shared-with-me-empty"]').exists()).toBe(true);
  });

  it('lists shared entries with name and permission', async () => {
    mockListSharedWithMe.mockResolvedValue([sharedEntry]);
    const wrapper = mountView();
    await flushPromises();

    const items = wrapper.findAll('[data-testid="office-shared-with-me-item"]');
    expect(items).toHaveLength(1);
    expect(wrapper.find('[data-testid="office-shared-with-me-name"]').text()).toBe('Roadmap.pdf');
    expect(wrapper.find('[data-testid="office-shared-with-me-permission"]').text()).toContain(
      'office.share.permission.view',
    );
  });

  it('opens a shared document via the authed download route', async () => {
    mockListSharedWithMe.mockResolvedValue([sharedEntry]);
    mockDownloadSharedDocument.mockResolvedValue({
      blob: new Blob(['content']),
      filename: 'Roadmap.pdf',
    });
    const wrapper = mountView();
    await flushPromises();

    await wrapper.find('[data-testid="office-shared-with-me-open"]').trigger('click');
    await flushPromises();

    expect(mockDownloadSharedDocument).toHaveBeenCalledWith('d1', 'Roadmap.pdf');
  });
});
