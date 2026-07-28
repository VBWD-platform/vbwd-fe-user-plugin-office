import { defineStore } from 'pinia';
import { ref } from 'vue';
import {
  officeApi,
  type OfficeNode,
  type OfficeUsage,
  type OfficeVersion,
} from '../api/officeApi';

export interface OfficeBreadcrumbEntry {
  id: string | null;
  name: string;
}

export type OfficeUploadStatus = 'uploading' | 'done' | 'error';

export interface OfficeUploadItem {
  id: string;
  name: string;
  progress: number;
  status: OfficeUploadStatus;
  errorMessage?: string;
}

let uploadItemSequence = 0;

export const useOfficeStore = defineStore('office', () => {
  const currentParentId = ref<string | null>(null);
  const breadcrumbs = ref<OfficeBreadcrumbEntry[]>([{ id: null, name: 'root' }]);
  const nodes = ref<OfficeNode[]>([]);
  const loading = ref(false);
  const error = ref<string | null>(null);
  const viewMode = ref<'list' | 'grid'>('list');
  const usage = ref<OfficeUsage | null>(null);
  const uploads = ref<OfficeUploadItem[]>([]);
  const selectedNode = ref<OfficeNode | null>(null);
  const versions = ref<OfficeVersion[]>([]);

  async function fetchNodes(parentId: string | null = currentParentId.value): Promise<void> {
    loading.value = true;
    error.value = null;
    try {
      nodes.value = await officeApi.listNodes(parentId);
    } catch (caught) {
      error.value = (caught as Error).message;
      nodes.value = [];
    } finally {
      loading.value = false;
    }
  }

  async function fetchUsage(): Promise<void> {
    try {
      usage.value = await officeApi.getUsage();
    } catch {
      // Usage is a secondary indicator — a failed fetch keeps the bar hidden
      // rather than surfacing a second error banner over the file listing.
      usage.value = null;
    }
  }

  /** Navigate into a folder: push a breadcrumb entry and refresh the listing. */
  async function openFolder(node: OfficeNode): Promise<void> {
    breadcrumbs.value = [...breadcrumbs.value, { id: node.id, name: node.name }];
    currentParentId.value = node.id;
    await fetchNodes(node.id);
  }

  /** Jump back to a prior breadcrumb, truncating everything after it. */
  async function goToBreadcrumb(index: number): Promise<void> {
    breadcrumbs.value = breadcrumbs.value.slice(0, index + 1);
    const target = breadcrumbs.value[breadcrumbs.value.length - 1];
    currentParentId.value = target.id;
    await fetchNodes(target.id);
  }

  /**
   * Jump straight to a folder from the sidebar tree, resetting the breadcrumb
   * trail to root → folder. The tree does not track full ancestry, so an
   * intermediate-level jump collapses the trail to these two entries — a
   * deliberate v1 simplification; the folder tree itself still shows the
   * real hierarchy.
   */
  async function navigateToNode(node: OfficeNode): Promise<void> {
    breadcrumbs.value = [{ id: null, name: 'root' }, { id: node.id, name: node.name }];
    currentParentId.value = node.id;
    await fetchNodes(node.id);
  }

  async function createFolder(name: string): Promise<void> {
    await officeApi.createFolder(name, currentParentId.value);
    await fetchNodes();
  }

  /**
   * Patch one tracked upload item immutably. Mutating a plain object captured
   * in a closure would bypass Vue's reactive proxy (the proxy wraps a COPY of
   * what was pushed, not the closed-over reference), so progress/status
   * updates always go through a fresh array replacement here.
   */
  function patchUploadItem(uploadId: string, patch: Partial<OfficeUploadItem>): void {
    uploads.value = uploads.value.map((item) =>
      item.id === uploadId ? { ...item, ...patch } : item,
    );
  }

  /**
   * Upload one or more files into the current folder, tracking a progress
   * item per file so the drag-drop dropzone can render a live indicator.
   * Failures are recorded on the item rather than thrown, so one bad file in
   * a multi-file drop does not abort the rest.
   */
  async function uploadFiles(files: File[]): Promise<void> {
    const parentId = currentParentId.value;
    await Promise.all(
      files.map(async (file) => {
        const uploadId = `upload-${++uploadItemSequence}`;
        uploads.value = [
          ...uploads.value,
          { id: uploadId, name: file.name, progress: 0, status: 'uploading' },
        ];

        try {
          await officeApi.uploadDocument(file, parentId, (fractionComplete) => {
            patchUploadItem(uploadId, { progress: fractionComplete });
          });
          patchUploadItem(uploadId, { status: 'done', progress: 1 });
          await fetchNodes();
          await fetchUsage();
        } catch (caught) {
          patchUploadItem(uploadId, { status: 'error', errorMessage: (caught as Error).message });
        }
      }),
    );
  }

  function dismissUpload(uploadId: string): void {
    uploads.value = uploads.value.filter((item) => item.id !== uploadId);
  }

  async function renameNode(node: OfficeNode, name: string): Promise<void> {
    await officeApi.renameNode(node.id, name);
    await fetchNodes();
  }

  async function moveNode(node: OfficeNode, parentId: string | null): Promise<void> {
    await officeApi.moveNode(node.id, parentId);
    await fetchNodes();
  }

  async function trashNode(node: OfficeNode, purge = false): Promise<void> {
    await officeApi.trashNode(node.id, purge);
    await fetchNodes();
    await fetchUsage();
  }

  async function fetchVersions(node: OfficeNode): Promise<void> {
    versions.value = await officeApi.listVersions(node.id);
  }

  /** `versionNo` — the version's ORDINAL (`version_no`), not its row id; the
   * backend's restore contract is `{version_no: <int>}`. */
  async function restoreVersion(node: OfficeNode, versionNo: number): Promise<void> {
    await officeApi.restoreVersion(node.id, versionNo);
    await fetchVersions(node);
    await fetchNodes();
  }

  function toggleViewMode(): void {
    viewMode.value = viewMode.value === 'list' ? 'grid' : 'list';
  }

  function selectNode(node: OfficeNode | null): void {
    selectedNode.value = node;
  }

  return {
    currentParentId,
    breadcrumbs,
    nodes,
    loading,
    error,
    viewMode,
    usage,
    uploads,
    selectedNode,
    versions,
    fetchNodes,
    fetchUsage,
    openFolder,
    goToBreadcrumb,
    navigateToNode,
    createFolder,
    uploadFiles,
    dismissUpload,
    renameNode,
    moveNode,
    trashNode,
    fetchVersions,
    restoreVersion,
    toggleViewMode,
    selectNode,
  };
});
