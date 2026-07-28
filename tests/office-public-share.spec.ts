import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';

const mockGetMetadata = vi.fn();
const mockUnlock = vi.fn();
const mockDownload = vi.fn();
const mockReplaceContent = vi.fn();

vi.mock('../src/api/officeApi', () => {
  class FakeUnavailableError extends Error {}
  class FakePasswordRequiredError extends Error {}
  return {
    officePublicShareApi: {
      getMetadata: (...args: unknown[]) => mockGetMetadata(...args),
      unlock: (...args: unknown[]) => mockUnlock(...args),
      download: (...args: unknown[]) => mockDownload(...args),
      replaceContent: (...args: unknown[]) => mockReplaceContent(...args),
    },
    OfficePublicShareUnavailableError: FakeUnavailableError,
    OfficePublicPasswordRequiredError: FakePasswordRequiredError,
  };
});

import OfficePublicShare from '../src/views/OfficePublicShare.vue';
import {
  OfficePublicShareUnavailableError as FakeUnavailableError,
} from '../src/api/officeApi';

const tFallback = (key: string) => key;

function mountView(token = 'vbwds_abc123') {
  return mount(OfficePublicShare, {
    props: { token },
    global: { mocks: { $t: tFallback } },
  });
}

describe('OfficePublicShare — the /s/:token page a stranger sees', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    // jsdom does not implement blob object URLs; the click-to-download path
    // only needs these to exist, not do anything real.
    global.URL.createObjectURL = vi.fn(() => 'blob:mock');
    global.URL.revokeObjectURL = vi.fn();
  });

  it('shows a loading state while metadata resolves', async () => {
    let resolveMetadata: (value: unknown) => void = () => {};
    mockGetMetadata.mockReturnValue(new Promise((resolve) => { resolveMetadata = resolve; }));

    const wrapper = mountView();
    expect(wrapper.find('[data-testid="office-public-loading"]').exists()).toBe(true);

    resolveMetadata({ name: 'x', kind: 'document', permission: 'view', requires_password: false });
    await flushPromises();
    expect(wrapper.find('[data-testid="office-public-loading"]').exists()).toBe(false);
  });

  it('shows a clean unavailable message for an invalid/revoked/expired token — never a crash', async () => {
    mockGetMetadata.mockRejectedValue(new FakeUnavailableError());
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-public-error"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-public-name"]').exists()).toBe(false);
  });

  it('renders the document name and a download button for a valid view link', async () => {
    mockGetMetadata.mockResolvedValue({
      name: 'Report.pdf',
      kind: 'document',
      permission: 'view',
      requires_password: false,
      document: { mime_type: 'application/pdf', size_bytes: 1024 },
    });
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-public-name"]').text()).toBe('Report.pdf');
    expect(wrapper.find('[data-testid="office-public-download-button"]').exists()).toBe(true);
    // View permission — no replace-file control.
    expect(wrapper.find('[data-testid="office-public-replace-input"]').exists()).toBe(false);
  });

  it('shows a password form when the share requires one, and unlocks it', async () => {
    mockGetMetadata.mockResolvedValue({
      name: 'Secret.txt',
      kind: 'document',
      permission: 'view',
      requires_password: true,
    });
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-public-password-form"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-public-download-button"]').exists()).toBe(false);

    mockUnlock.mockResolvedValue('grant-token-123');
    await wrapper.find('[data-testid="office-public-password-input"]').setValue('hunter2');
    await wrapper.find('[data-testid="office-public-password-form"]').trigger('submit');
    await flushPromises();

    expect(mockUnlock).toHaveBeenCalledWith('vbwds_abc123', 'hunter2');
    expect(wrapper.find('[data-testid="office-public-password-form"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="office-public-download-button"]').exists()).toBe(true);
  });

  it('shows an error and does not unlock on a wrong password', async () => {
    mockGetMetadata.mockResolvedValue({
      name: 'Secret.txt',
      kind: 'document',
      permission: 'view',
      requires_password: true,
    });
    mockUnlock.mockRejectedValue(new Error('Invalid password'));
    const wrapper = mountView();
    await flushPromises();

    await wrapper.find('[data-testid="office-public-password-input"]').setValue('wrong');
    await wrapper.find('[data-testid="office-public-password-form"]').trigger('submit');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-public-password-error"]').text()).toContain(
      'Invalid password',
    );
    expect(wrapper.find('[data-testid="office-public-password-form"]').exists()).toBe(true);
  });

  it('triggers a download carrying no auth header via officePublicShareApi', async () => {
    mockGetMetadata.mockResolvedValue({
      name: 'Report.pdf',
      kind: 'document',
      permission: 'view',
      requires_password: false,
    });
    mockDownload.mockResolvedValue({ blob: new Blob(['x']), filename: 'Report.pdf' });
    const wrapper = mountView();
    await flushPromises();

    await wrapper.find('[data-testid="office-public-download-button"]').trigger('click');
    await flushPromises();

    expect(mockDownload).toHaveBeenCalledWith('vbwds_abc123', 'Report.pdf', null);
  });

  it('shows a replace-file control only when permission is edit', async () => {
    mockGetMetadata.mockResolvedValue({
      name: 'Notes.txt',
      kind: 'document',
      permission: 'edit',
      requires_password: false,
    });
    const wrapper = mountView();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-public-replace-input"]').exists()).toBe(true);
  });
});
