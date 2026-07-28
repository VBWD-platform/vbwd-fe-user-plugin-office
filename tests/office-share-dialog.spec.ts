import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';

const mockListShares = vi.fn();
const mockCreateShare = vi.fn();
const mockRevokeShare = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    listShares: (...args: unknown[]) => mockListShares(...args),
    createShare: (...args: unknown[]) => mockCreateShare(...args),
    revokeShare: (...args: unknown[]) => mockRevokeShare(...args),
  },
}));

import OfficeShareDialog from '../src/components/OfficeShareDialog.vue';

const tFallback = (key: string, params?: Record<string, unknown>) =>
  params ? `${key}:${JSON.stringify(params)}` : key;

const documentNode = {
  id: 'd1',
  parent_id: null,
  kind: 'document' as const,
  name: 'Report.pdf',
  trashed_at: null,
  created_at: '2026-07-27T00:00:00Z',
  updated_at: '2026-07-27T00:00:00Z',
};

function mountDialog() {
  return mount(OfficeShareDialog, {
    props: { node: documentNode },
    global: { mocks: { $t: tFallback } },
  });
}

describe('OfficeShareDialog', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockListShares.mockResolvedValue([]);
  });

  it('loads and lists existing shares on mount', async () => {
    mockListShares.mockResolvedValue([
      {
        id: 's1',
        node_id: 'd1',
        created_by_user_id: 'u1',
        permission: 'view',
        subject_user_id: null,
        allow_anonymous: true,
        has_password: false,
        expires_at: null,
        revoked_at: null,
        last_used_at: null,
        created_at: '2026-07-27T00:00:00Z',
        updated_at: '2026-07-27T00:00:00Z',
      },
    ]);
    const wrapper = mountDialog();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-share-list"]').exists()).toBe(true);
    expect(wrapper.findAll('[data-testid="office-share-item"]')).toHaveLength(1);
  });

  it('shows an empty state when nothing has been shared yet', async () => {
    const wrapper = mountDialog();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-share-empty"]').exists()).toBe(true);
  });

  it('creates a share and shows the token exactly once, with a copy button', async () => {
    mockCreateShare.mockResolvedValue({
      id: 's2',
      node_id: 'd1',
      created_by_user_id: 'u1',
      permission: 'view',
      subject_user_id: null,
      allow_anonymous: true,
      has_password: false,
      expires_at: null,
      revoked_at: null,
      last_used_at: null,
      created_at: '2026-07-27T00:00:00Z',
      updated_at: '2026-07-27T00:00:00Z',
      token: 'vbwds_secret-token',
    });
    const wrapper = mountDialog();
    await flushPromises();

    await wrapper.find('[data-testid="office-share-permission-edit"]').setValue(true);
    await wrapper.find('[data-testid="office-share-anonymous-toggle"]').setValue(true);
    mockListShares.mockResolvedValue([{ id: 's2', permission: 'edit' }]);
    await wrapper.find('[data-testid="office-share-form"]').trigger('submit');
    await flushPromises();

    expect(mockCreateShare).toHaveBeenCalledWith(
      'd1',
      expect.objectContaining({ permission: 'edit', allow_anonymous: true }),
    );

    const tokenPanel = wrapper.find('[data-testid="office-share-token-display"]');
    expect(tokenPanel.exists()).toBe(true);
    const linkInput = wrapper.find('[data-testid="office-share-link-input"]');
    expect((linkInput.element as HTMLInputElement).value).toContain('vbwds_secret-token');
    expect(wrapper.find('[data-testid="office-share-copy-button"]').exists()).toBe(true);
  });

  it('shows a create error inline without crashing', async () => {
    mockCreateShare.mockRejectedValue(new Error('permission must be one of...'));
    const wrapper = mountDialog();
    await flushPromises();

    await wrapper.find('[data-testid="office-share-form"]').trigger('submit');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-share-error"]').text()).toContain(
      'permission must be one of',
    );
  });

  it('revokes a share and removes it from the list', async () => {
    mockListShares.mockResolvedValueOnce([
      {
        id: 's1',
        node_id: 'd1',
        created_by_user_id: 'u1',
        permission: 'view',
        subject_user_id: null,
        allow_anonymous: true,
        has_password: false,
        expires_at: null,
        revoked_at: null,
        last_used_at: null,
        created_at: '2026-07-27T00:00:00Z',
        updated_at: '2026-07-27T00:00:00Z',
      },
    ]);
    const wrapper = mountDialog();
    await flushPromises();
    expect(wrapper.findAll('[data-testid="office-share-item"]')).toHaveLength(1);

    mockRevokeShare.mockResolvedValue(undefined);
    mockListShares.mockResolvedValueOnce([]);
    await wrapper.find('[data-testid="office-share-revoke-button"]').trigger('click');
    await flushPromises();

    expect(mockRevokeShare).toHaveBeenCalledWith('s1');
    expect(wrapper.find('[data-testid="office-share-empty"]').exists()).toBe(true);
  });

  it('emits close when the close button is clicked', async () => {
    const wrapper = mountDialog();
    await flushPromises();

    await wrapper.find('[data-testid="office-share-close-button"]').trigger('click');

    expect(wrapper.emitted('close')).toBeTruthy();
  });
});
