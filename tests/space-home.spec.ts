import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises, RouterLinkStub } from '@vue/test-utils';
import { createPinia, setActivePinia } from 'pinia';

const mockListNodes = vi.fn();
const mockGetUsage = vi.fn();
const mockCreateFolder = vi.fn();
const mockUploadDocument = vi.fn();
const mockDownloadDocument = vi.fn();
const mockListVersions = vi.fn();
const mockCreateDoc = vi.fn();
const mockCreateSheet = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    listNodes: (...args: unknown[]) => mockListNodes(...args),
    getUsage: (...args: unknown[]) => mockGetUsage(...args),
    createFolder: (...args: unknown[]) => mockCreateFolder(...args),
    uploadDocument: (...args: unknown[]) => mockUploadDocument(...args),
    downloadDocument: (...args: unknown[]) => mockDownloadDocument(...args),
    listVersions: (...args: unknown[]) => mockListVersions(...args),
    renameNode: vi.fn(),
    moveNode: vi.fn(),
    trashNode: vi.fn(),
    restoreVersion: vi.fn(),
    contentUrl: (id: string) => `/api/v1/office/documents/${id}/content`,
  },
  officeDocApi: {
    createDoc: (...args: unknown[]) => mockCreateDoc(...args),
  },
  officeSheetApi: {
    createSheet: (...args: unknown[]) => mockCreateSheet(...args),
  },
}));

const mockRouterPush = vi.fn();
vi.mock('vue-router', () => ({
  useRouter: () => ({ push: mockRouterPush }),
}));

// Stub the shared Icon component so mounting does not need the fe-core registry.
vi.mock('vbwd-view-component', () => ({
  Icon: { name: 'IconStub', props: ['name', 'size'], template: '<span :data-icon="name" />' },
}));

import SpaceHome from '../src/views/SpaceHome.vue';

const tFallback = (key: string) => key;

function mountSpaceHome() {
  return mount(SpaceHome, {
    global: {
      mocks: { $t: tFallback },
      stubs: { RouterLink: RouterLinkStub },
    },
  });
}

const documentNode = {
  id: 'd1',
  parent_id: null,
  kind: 'document' as const,
  name: 'Report.pdf',
  trashed_at: null,
  created_at: '2026-07-27T00:00:00Z',
  updated_at: '2026-07-27T00:00:00Z',
  mime_type: 'application/pdf',
  size_bytes: 2048,
};

const folderNode = {
  id: 'f1',
  parent_id: null,
  kind: 'folder' as const,
  name: 'Contracts',
  trashed_at: null,
  created_at: '2026-07-27T00:00:00Z',
  updated_at: '2026-07-27T00:00:00Z',
};

describe('SpaceHome', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    mockGetUsage.mockResolvedValue({ bytes_used: 100, bytes_quota: 1000 });
  });

  it('shows the loading state while nodes are being fetched', async () => {
    let resolveNodes: (value: unknown[]) => void = () => {};
    mockListNodes.mockReturnValue(new Promise((resolve) => { resolveNodes = resolve; }));

    const wrapper = mountSpaceHome();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-loading"]').exists()).toBe(true);

    resolveNodes([]);
    await flushPromises();
    expect(wrapper.find('[data-testid="office-loading"]').exists()).toBe(false);
  });

  it('shows a genuinely helpful empty state with an upload call-to-action', async () => {
    mockListNodes.mockResolvedValue([]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-empty-state"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-empty-upload-cta"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-node-list"]').exists()).toBe(false);
  });

  it('renders populated folders and documents with stable testids', async () => {
    mockListNodes.mockResolvedValue([folderNode, documentNode]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    const items = wrapper.findAll('[data-testid="office-node-item"]');
    expect(items).toHaveLength(2);
    const names = wrapper.findAll('[data-testid="office-node-name"]').map((n) => n.text());
    expect(names).toContain('Contracts');
    expect(names).toContain('Report.pdf');
  });

  it('shows the error banner with a working retry action', async () => {
    mockListNodes.mockRejectedValueOnce(new Error('GET failed: 500'));
    const wrapper = mountSpaceHome();
    await flushPromises();

    const banner = wrapper.find('[data-testid="office-error-banner"]');
    expect(banner.exists()).toBe(true);
    expect(banner.text()).toContain('GET failed: 500');

    mockListNodes.mockResolvedValueOnce([folderNode]);
    await wrapper.find('[data-testid="office-error-retry"]').trigger('click');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-error-banner"]').exists()).toBe(false);
    expect(wrapper.findAll('[data-testid="office-node-item"]')).toHaveLength(1);
  });

  it('shows an uploading progress item while a drag-drop upload is in flight, then done', async () => {
    mockListNodes.mockResolvedValue([]);
    let resolveUpload: (value: unknown) => void = () => {};
    mockUploadDocument.mockImplementation(
      (_file: File, _parentId: string | null, onProgress?: (fraction: number) => void) =>
        new Promise((resolve) => {
          onProgress?.(0.4);
          resolveUpload = () => resolve({ ...documentNode, id: 'new-doc' });
        }),
    );

    const wrapper = mountSpaceHome();
    await flushPromises();

    const file = new File(['binary'], 'upload.pdf', { type: 'application/pdf' });
    const dataTransfer = { files: [file] } as unknown as DataTransfer;
    await wrapper.find('[data-testid="office-content"]').trigger('drop', { dataTransfer });
    await flushPromises();

    const progressItem = wrapper.find('[data-testid="office-upload-progress-item"]');
    expect(progressItem.exists()).toBe(true);
    expect(progressItem.text()).toContain('upload.pdf');
    expect(wrapper.find('[data-testid="office-upload-progress-bar"]').exists()).toBe(true);

    mockListNodes.mockResolvedValue([documentNode]);
    resolveUpload({});
    await flushPromises();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-upload-progress-item"]').classes()).toContain('is-done');
  });

  it('shows the drag-active dropzone overlay while dragging over the content area', async () => {
    mockListNodes.mockResolvedValue([]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-content"]').trigger('dragover');
    expect(wrapper.find('[data-testid="office-dropzone"]').exists()).toBe(true);

    await wrapper.find('[data-testid="office-content"]').trigger('dragleave');
    expect(wrapper.find('[data-testid="office-dropzone"]').exists()).toBe(false);
  });

  it('creates a folder via the new-folder dialog', async () => {
    mockListNodes.mockResolvedValue([]);
    mockCreateFolder.mockResolvedValue(folderNode);
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-new-folder-btn"]').trigger('click');
    expect(wrapper.find('[data-testid="office-new-folder-modal"]').exists()).toBe(true);

    await wrapper.find('[data-testid="office-new-folder-input"]').setValue('Contracts');
    mockListNodes.mockResolvedValue([folderNode]);
    await wrapper.find('[data-testid="office-new-folder-confirm"]').trigger('click');
    await flushPromises();

    expect(mockCreateFolder).toHaveBeenCalledWith('Contracts', null);
    expect(wrapper.find('[data-testid="office-new-folder-modal"]').exists()).toBe(false);
  });

  it('creates a document via the new-document dialog', async () => {
    mockListNodes.mockResolvedValue([]);
    mockCreateDoc.mockResolvedValue({ id: 'doc-1' });
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-new-doc-btn"]').trigger('click');
    expect(wrapper.find('[data-testid="office-new-doc-modal"]').exists()).toBe(true);

    await wrapper.find('[data-testid="office-new-doc-input"]').setValue('Meeting notes');
    await wrapper.find('[data-testid="office-new-doc-confirm"]').trigger('click');
    await flushPromises();

    expect(mockCreateDoc).toHaveBeenCalledWith('Meeting notes', null);
    expect(mockRouterPush).toHaveBeenCalledWith('/dashboard/office/doc/doc-1');
    expect(wrapper.find('[data-testid="office-new-doc-modal"]').exists()).toBe(false);
  });

  it('cancelling the new-document dialog creates nothing', async () => {
    mockListNodes.mockResolvedValue([]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-new-doc-btn"]').trigger('click');
    await wrapper.find('[data-testid="office-new-doc-cancel"]').trigger('click');

    expect(mockCreateDoc).not.toHaveBeenCalled();
    expect(wrapper.find('[data-testid="office-new-doc-modal"]').exists()).toBe(false);
  });

  it('creates a spreadsheet via the new-spreadsheet dialog', async () => {
    mockListNodes.mockResolvedValue([]);
    mockCreateSheet.mockResolvedValue({ id: 'sheet-1' });
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-new-sheet-btn"]').trigger('click');
    expect(wrapper.find('[data-testid="office-new-sheet-modal"]').exists()).toBe(true);

    await wrapper.find('[data-testid="office-new-sheet-input"]').setValue('Budget');
    await wrapper.find('[data-testid="office-new-sheet-confirm"]').trigger('click');
    await flushPromises();

    expect(mockCreateSheet).toHaveBeenCalledWith('Budget', null);
    expect(mockRouterPush).toHaveBeenCalledWith('/dashboard/office/sheet/sheet-1');
    expect(wrapper.find('[data-testid="office-new-sheet-modal"]').exists()).toBe(false);
  });

  it('cancelling the new-spreadsheet dialog creates nothing', async () => {
    mockListNodes.mockResolvedValue([]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-new-sheet-btn"]').trigger('click');
    await wrapper.find('[data-testid="office-new-sheet-cancel"]').trigger('click');

    expect(mockCreateSheet).not.toHaveBeenCalled();
    expect(wrapper.find('[data-testid="office-new-sheet-modal"]').exists()).toBe(false);
  });

  it('opens the preview pane when a document is clicked', async () => {
    mockListNodes.mockResolvedValue([documentNode]);
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-node-item"]').trigger('click');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-preview-pane"]').exists()).toBe(true);
  });
});


/**
 * Coverage for the context-menu wiring: visibility is a SEPARATE flag from the
 * target node, because `node: null` is a meaningful state (right-click on empty
 * space offers Paste alone) and therefore cannot double as "closed". Getting that
 * wrong either hides the empty-space menu or leaves a stale menu open.
 */
describe('SpaceHome — right-click context menu', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    mockGetUsage.mockResolvedValue({ bytes_used: 0, bytes_quota: 1024 });
    mockListNodes.mockResolvedValue([documentNode, folderNode]);
  });

  it('right-clicking a row opens the menu targeting that node', async () => {
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.findAll('[data-testid="office-node-item"]')[0].trigger('contextmenu');
    await flushPromises();

    const menu = wrapper.find('[data-testid="office-context-menu"]');
    expect(menu.exists()).toBe(true);
    // A node-targeted menu offers the node actions, not just Paste.
    expect(wrapper.find('[data-testid="office-context-copy"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-context-cut"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-context-delete"]').exists()).toBe(true);
  });

  it('right-clicking empty space offers Paste only, disabled while the clipboard is empty', async () => {
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.find('[data-testid="office-content"]').trigger('contextmenu');
    await flushPromises();

    const paste = wrapper.find('[data-testid="office-context-paste"]');
    expect(paste.exists()).toBe(true);
    expect(paste.attributes('disabled')).toBeDefined();
    // No node actions on an empty-space menu.
    expect(wrapper.find('[data-testid="office-context-copy"]').exists()).toBe(false);
  });

  it('enables Paste once something is on the clipboard', async () => {
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.findAll('[data-testid="office-node-item"]')[0].trigger('contextmenu');
    await flushPromises();
    await wrapper.find('[data-testid="office-context-copy"]').trigger('click');
    await flushPromises();

    await wrapper.find('[data-testid="office-content"]').trigger('contextmenu');
    await flushPromises();

    expect(
      wrapper.find('[data-testid="office-context-paste"]').attributes('disabled'),
    ).toBeUndefined();
  });

  it('Delete asks for confirmation instead of purging straight away', async () => {
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.findAll('[data-testid="office-node-item"]')[0].trigger('contextmenu');
    await flushPromises();
    await wrapper.find('[data-testid="office-context-delete"]').trigger('click');
    await flushPromises();

    // The confirm dialog is the gate — a destructive action must never fire off a
    // single menu click.
    expect(wrapper.find('[data-testid="office-delete-modal"]').exists()).toBe(true);
  });

  it('cancelling the delete confirmation removes nothing', async () => {
    const wrapper = mountSpaceHome();
    await flushPromises();

    await wrapper.findAll('[data-testid="office-node-item"]')[0].trigger('contextmenu');
    await flushPromises();
    await wrapper.find('[data-testid="office-context-delete"]').trigger('click');
    await flushPromises();
    await wrapper.find('[data-testid="office-delete-cancel"]').trigger('click');
    await flushPromises();

    expect(wrapper.find('[data-testid="office-delete-modal"]').exists()).toBe(false);
    expect(wrapper.findAll('[data-testid="office-node-item"]')).toHaveLength(2);
  });
});
