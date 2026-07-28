import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';

const mockListNodes = vi.fn();
const mockCreateFolder = vi.fn();
const mockUploadDocument = vi.fn();
const mockRenameNode = vi.fn();
const mockMoveNode = vi.fn();
const mockTrashNode = vi.fn();
const mockGetUsage = vi.fn();
const mockListVersions = vi.fn();
const mockRestoreVersion = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    listNodes: (...args: unknown[]) => mockListNodes(...args),
    createFolder: (...args: unknown[]) => mockCreateFolder(...args),
    uploadDocument: (...args: unknown[]) => mockUploadDocument(...args),
    renameNode: (...args: unknown[]) => mockRenameNode(...args),
    moveNode: (...args: unknown[]) => mockMoveNode(...args),
    trashNode: (...args: unknown[]) => mockTrashNode(...args),
    getUsage: (...args: unknown[]) => mockGetUsage(...args),
    listVersions: (...args: unknown[]) => mockListVersions(...args),
    restoreVersion: (...args: unknown[]) => mockRestoreVersion(...args),
  },
}));

import { useOfficeStore } from '../src/stores/useOfficeStore';

const folderNode = {
  id: 'f1',
  parent_id: null,
  kind: 'folder' as const,
  name: 'Reports',
  trashed_at: null,
  created_at: '2026-07-27T00:00:00Z',
  updated_at: '2026-07-27T00:00:00Z',
};

describe('useOfficeStore', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
  });

  it('starts at the root with an empty breadcrumb trail', () => {
    const store = useOfficeStore();
    expect(store.currentParentId).toBeNull();
    expect(store.breadcrumbs).toHaveLength(1);
    expect(store.breadcrumbs[0].id).toBeNull();
  });

  it('fetchNodes populates nodes and clears loading on success', async () => {
    mockListNodes.mockResolvedValue([folderNode]);
    const store = useOfficeStore();

    const promise = store.fetchNodes();
    expect(store.loading).toBe(true);
    await promise;

    expect(store.loading).toBe(false);
    expect(store.nodes).toEqual([folderNode]);
    expect(store.error).toBeNull();
  });

  it('fetchNodes records the error message and empties the listing on failure', async () => {
    mockListNodes.mockRejectedValue(new Error('GET failed: 500'));
    const store = useOfficeStore();

    await store.fetchNodes();

    expect(store.error).toBe('GET failed: 500');
    expect(store.nodes).toEqual([]);
    expect(store.loading).toBe(false);
  });

  it('openFolder pushes a breadcrumb and re-fetches scoped to the folder', async () => {
    mockListNodes.mockResolvedValue([]);
    const store = useOfficeStore();

    await store.openFolder(folderNode);

    expect(store.currentParentId).toBe('f1');
    expect(store.breadcrumbs).toHaveLength(2);
    expect(store.breadcrumbs[1]).toEqual({ id: 'f1', name: 'Reports' });
    expect(mockListNodes).toHaveBeenCalledWith('f1');
  });

  it('goToBreadcrumb truncates the trail and re-fetches at that level', async () => {
    mockListNodes.mockResolvedValue([]);
    const store = useOfficeStore();
    await store.openFolder(folderNode);
    await store.openFolder({ ...folderNode, id: 'f2', name: 'Nested' });
    expect(store.breadcrumbs).toHaveLength(3);

    await store.goToBreadcrumb(0);

    expect(store.breadcrumbs).toHaveLength(1);
    expect(store.currentParentId).toBeNull();
    expect(mockListNodes).toHaveBeenLastCalledWith(null);
  });

  it('createFolder calls the API scoped to the current folder then refreshes', async () => {
    mockCreateFolder.mockResolvedValue(folderNode);
    mockListNodes.mockResolvedValue([folderNode]);
    const store = useOfficeStore();

    await store.createFolder('Reports');

    expect(mockCreateFolder).toHaveBeenCalledWith('Reports', null);
    expect(store.nodes).toEqual([folderNode]);
  });

  it('uploadFiles tracks a progress item that reaches done on success', async () => {
    let capturedProgress: ((fraction: number) => void) | undefined;
    mockUploadDocument.mockImplementation(async (_file, _parentId, onProgress) => {
      capturedProgress = onProgress;
      onProgress?.(0.5);
      return { ...folderNode, kind: 'document', name: 'report.pdf' };
    });
    mockListNodes.mockResolvedValue([]);
    mockGetUsage.mockResolvedValue({ bytes_used: 100, bytes_quota: 1000 });
    const store = useOfficeStore();
    const file = new File(['data'], 'report.pdf', { type: 'application/pdf' });

    await store.uploadFiles([file]);

    expect(capturedProgress).toBeDefined();
    expect(store.uploads).toHaveLength(1);
    expect(store.uploads[0]).toMatchObject({ name: 'report.pdf', status: 'done', progress: 1 });
    expect(store.usage).toEqual({ bytes_used: 100, bytes_quota: 1000 });
  });

  it('uploadFiles records a failed item without throwing (one bad file does not abort the batch)', async () => {
    mockUploadDocument.mockRejectedValue(new Error('Upload failed: 413'));
    const store = useOfficeStore();
    const file = new File(['data'], 'huge.bin');

    await expect(store.uploadFiles([file])).resolves.toBeUndefined();

    expect(store.uploads[0]).toMatchObject({ status: 'error', errorMessage: 'Upload failed: 413' });
  });

  it('dismissUpload removes the item from the tracked list', async () => {
    mockUploadDocument.mockResolvedValue(folderNode);
    mockListNodes.mockResolvedValue([]);
    mockGetUsage.mockResolvedValue({ bytes_used: 0, bytes_quota: 1000 });
    const store = useOfficeStore();
    await store.uploadFiles([new File(['x'], 'a.txt')]);
    const uploadId = store.uploads[0].id;

    store.dismissUpload(uploadId);

    expect(store.uploads).toHaveLength(0);
  });

  it('renameNode, moveNode and trashNode each refresh the listing', async () => {
    mockRenameNode.mockResolvedValue(folderNode);
    mockMoveNode.mockResolvedValue(folderNode);
    mockTrashNode.mockResolvedValue(undefined);
    mockListNodes.mockResolvedValue([]);
    mockGetUsage.mockResolvedValue({ bytes_used: 0, bytes_quota: 1000 });
    const store = useOfficeStore();

    await store.renameNode(folderNode, 'Renamed');
    expect(mockRenameNode).toHaveBeenCalledWith('f1', 'Renamed');

    await store.moveNode(folderNode, 'f2');
    expect(mockMoveNode).toHaveBeenCalledWith('f1', 'f2');

    await store.trashNode(folderNode);
    expect(mockTrashNode).toHaveBeenCalledWith('f1', false);
    expect(mockGetUsage).toHaveBeenCalled();
  });

  it('fetchVersions and restoreVersion drive the version-history panel', async () => {
    mockListVersions.mockResolvedValue([{ id: 'v1', version_no: 1, size_bytes: 10, sha256: 'abc', created_by_user_id: 'u1', created_at: '2026-07-27T00:00:00Z' }]);
    mockRestoreVersion.mockResolvedValue(folderNode);
    mockListNodes.mockResolvedValue([]);
    const store = useOfficeStore();

    await store.fetchVersions(folderNode);
    expect(store.versions).toHaveLength(1);

    await store.restoreVersion(folderNode, 1);
    expect(mockRestoreVersion).toHaveBeenCalledWith('f1', 1);
  });

  it('toggleViewMode flips between list and grid', () => {
    const store = useOfficeStore();
    expect(store.viewMode).toBe('list');
    store.toggleViewMode();
    expect(store.viewMode).toBe('grid');
    store.toggleViewMode();
    expect(store.viewMode).toBe('list');
  });
});
