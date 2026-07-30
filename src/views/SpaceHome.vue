<template>
  <div
    class="office-space-home"
    data-testid="office-space-home"
  >
    <aside class="office-sidebar">
      <h2 class="vbwd-heading office-sidebar-title">
        {{ $t('office.title') }}
      </h2>
      <div
        class="office-sidebar-tree"
        data-testid="office-folder-tree"
      >
        <OfficeFolderTree
          :parent-id="null"
          :label="$t('office.root')"
          :depth="0"
          :active-id="store.currentParentId"
          @select="onTreeSelect"
        />
      </div>
      <OfficeUsageBar
        v-if="store.usage"
        :usage="store.usage"
      />
    </aside>

    <main class="office-main">
      <div
        class="office-toolbar"
        data-testid="office-toolbar"
      >
        <OfficeBreadcrumb
          :crumbs="store.breadcrumbs"
          @navigate="store.goToBreadcrumb"
        />
        <div class="office-toolbar-actions">
          <input
            v-model="searchTerm"
            type="search"
            class="vbwd-input"
            :placeholder="$t('office.toolbar.search')"
            data-testid="office-search-input"
          >
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-new-folder-btn"
            @click="showNewFolderDialog = true"
          >
            {{ $t('office.toolbar.newFolder') }}
          </button>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-new-doc-btn"
            @click="showNewDocDialog = true"
          >
            {{ $t('office.toolbar.newDoc') }}
          </button>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-new-sheet-btn"
            @click="showNewSheetDialog = true"
          >
            {{ $t('office.toolbar.newSheet') }}
          </button>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary vbwd-btn--sm"
            data-testid="office-upload-btn"
            @click="triggerFileInput"
          >
            {{ $t('office.toolbar.upload') }}
          </button>
          <input
            ref="fileInputRef"
            type="file"
            multiple
            class="office-hidden-input"
            data-testid="office-upload-input"
            @change="onFileInputChange"
          >
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-view-toggle"
            @click="store.toggleViewMode"
          >
            {{ store.viewMode === 'list' ? $t('office.toolbar.viewGrid') : $t('office.toolbar.viewList') }}
          </button>
          <RouterLink
            to="/dashboard/office/shared-with-me"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-shared-with-me-link"
          >
            {{ $t('office.toolbar.sharedWithMe') }}
          </RouterLink>
        </div>
      </div>

      <div
        class="office-content"
        :class="{ 'is-drag-active': isDragActive }"
        data-testid="office-content"
        @click="contextMenuNode = null"
        @dragover.prevent="isDragActive = true"
        @dragleave.prevent="isDragActive = false"
        @drop.prevent="onDrop"
      >
        <div
          v-if="isDragActive"
          class="office-dropzone-overlay"
          data-testid="office-dropzone"
        >
          <p>{{ $t('office.dropzone.prompt') }}</p>
        </div>

        <div
          v-if="store.loading"
          class="office-loading"
          data-testid="office-loading"
        >
          {{ $t('office.loading') }}
        </div>

        <div
          v-else-if="store.error"
          class="office-error-banner"
          data-testid="office-error-banner"
        >
          <p>{{ store.error }}</p>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
            data-testid="office-error-retry"
            @click="store.fetchNodes()"
          >
            {{ $t('office.retry') }}
          </button>
        </div>

        <div
          v-else-if="filteredNodes.length === 0"
          class="office-empty-state"
          data-testid="office-empty-state"
        >
          <Icon
            name="document"
            :size="40"
          />
          <p class="office-empty-title">
            {{ $t('office.empty.title') }}
          </p>
          <p class="office-empty-hint">
            {{ $t('office.empty.hint') }}
          </p>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary"
            data-testid="office-empty-upload-cta"
            @click="triggerFileInput"
          >
            {{ $t('office.empty.uploadCta') }}
          </button>
        </div>

        <ul
          v-else-if="store.viewMode === 'list'"
          class="office-node-list"
          data-testid="office-node-list"
        >
          <li
            v-for="node in filteredNodes"
            :key="node.id"
            class="office-node-row"
            data-testid="office-node-item"
            @click.stop="onNodeClick(node)"
          >
            <Icon :name="node.kind === 'folder' ? 'layers' : 'document'" />
            <span
              class="office-node-name"
              data-testid="office-node-name"
            >{{ node.name }}</span>
            <span class="office-node-modified">{{ formatDate(node.updated_at) }}</span>
            <span
              class="office-node-size"
              data-testid="office-node-size"
            >{{ node.kind === 'document' ? formatBytes(node.size_bytes ?? 0) : '—' }}</span>
            <button
              type="button"
              class="office-node-menu-btn"
              data-testid="office-node-menu-btn"
              @click.stop="onNodeMenu(node, $event)"
            >
              ⋮
            </button>
          </li>
        </ul>

        <div
          v-else
          class="office-node-grid"
          data-testid="office-node-grid"
        >
          <div
            v-for="node in filteredNodes"
            :key="node.id"
            class="office-node-tile"
            data-testid="office-node-item"
            @click.stop="onNodeClick(node)"
          >
            <Icon
              :name="node.kind === 'folder' ? 'layers' : 'document'"
              :size="32"
            />
            <span
              class="office-node-name"
              data-testid="office-node-name"
            >{{ node.name }}</span>
            <button
              type="button"
              class="office-node-menu-btn"
              data-testid="office-node-menu-btn"
              @click.stop="onNodeMenu(node, $event)"
            >
              ⋮
            </button>
          </div>
        </div>
      </div>

      <OfficeContextMenu
        v-if="contextMenuNode"
        :node="contextMenuNode"
        :x="contextMenuX"
        :y="contextMenuY"
        @action="onContextAction"
      />
    </main>

    <OfficePreviewPane
      v-if="previewNode"
      :node="previewNode"
      @close="previewNode = null"
    />

    <OfficeInputDialog
      v-if="showNewFolderDialog"
      :title="$t('office.newFolder.title')"
      :placeholder="$t('office.newFolder.placeholder')"
      :confirm-label="$t('office.newFolder.create')"
      :cancel-label="$t('office.newFolder.cancel')"
      testid-prefix="office-new-folder"
      @confirm="onCreateFolder"
      @cancel="showNewFolderDialog = false"
    />

    <OfficeInputDialog
      v-if="showNewDocDialog"
      :title="$t('office.newDoc.title')"
      :placeholder="$t('office.newDoc.namePrompt')"
      :confirm-label="$t('office.newDoc.create')"
      :cancel-label="$t('office.newDoc.cancel')"
      initial-value="Untitled document"
      testid-prefix="office-new-doc"
      @confirm="onCreateDoc"
      @cancel="showNewDocDialog = false"
    />

    <OfficeInputDialog
      v-if="showNewSheetDialog"
      :title="$t('office.newSheet.title')"
      :placeholder="$t('office.newSheet.namePrompt')"
      :confirm-label="$t('office.newSheet.create')"
      :cancel-label="$t('office.newSheet.cancel')"
      initial-value="Untitled spreadsheet"
      testid-prefix="office-new-sheet"
      @confirm="onCreateSheet"
      @cancel="showNewSheetDialog = false"
    />

    <OfficeInputDialog
      v-if="renamingNode"
      :title="$t('office.rename.title')"
      :placeholder="$t('office.rename.placeholder')"
      :confirm-label="$t('office.rename.save')"
      :cancel-label="$t('office.rename.cancel')"
      :initial-value="renamingNode.name"
      testid-prefix="office-rename"
      @confirm="onRenameConfirm"
      @cancel="renamingNode = null"
    />

    <OfficeMoveDialog
      v-if="movingNode"
      @confirm="onMoveConfirm"
      @cancel="movingNode = null"
    />

    <OfficeVersionsDialog
      v-if="versionsNode"
      :versions="store.versions"
      @close="versionsNode = null"
      @restore="onRestoreVersion"
    />

    <OfficeShareDialog
      v-if="sharingNode"
      :node="sharingNode"
      @close="sharingNode = null"
    />

    <OfficeUploadProgressList
      :uploads="store.uploads"
      @dismiss="store.dismissUpload"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import { useRouter } from 'vue-router';
import { Icon } from 'vbwd-view-component';
import { officeApi, officeDocApi, officeSheetApi, type OfficeNode } from '../api/officeApi';
import { useOfficeStore } from '../stores/useOfficeStore';
import OfficeBreadcrumb from '../components/OfficeBreadcrumb.vue';
import OfficeFolderTree from '../components/OfficeFolderTree.vue';
import OfficeContextMenu, { type OfficeContextMenuAction } from '../components/OfficeContextMenu.vue';
import OfficeUsageBar from '../components/OfficeUsageBar.vue';
import OfficePreviewPane from '../components/OfficePreviewPane.vue';
import OfficeInputDialog from '../components/OfficeInputDialog.vue';
import OfficeMoveDialog from '../components/OfficeMoveDialog.vue';
import OfficeVersionsDialog from '../components/OfficeVersionsDialog.vue';
import OfficeShareDialog from '../components/OfficeShareDialog.vue';
import OfficeUploadProgressList from '../components/OfficeUploadProgressList.vue';

const store = useOfficeStore();
const router = useRouter();

const searchTerm = ref('');
const isDragActive = ref(false);
const fileInputRef = ref<HTMLInputElement | null>(null);

const previewNode = ref<OfficeNode | null>(null);
const contextMenuNode = ref<OfficeNode | null>(null);
const contextMenuX = ref(0);
const contextMenuY = ref(0);

const showNewFolderDialog = ref(false);
const showNewDocDialog = ref(false);
const showNewSheetDialog = ref(false);
const renamingNode = ref<OfficeNode | null>(null);
const movingNode = ref<OfficeNode | null>(null);
const versionsNode = ref<OfficeNode | null>(null);
const sharingNode = ref<OfficeNode | null>(null);

const filteredNodes = computed(() => {
  const term = searchTerm.value.trim().toLowerCase();
  const activeNodes = store.nodes.filter((node) => !node.trashed_at);
  if (!term) return activeNodes;
  return activeNodes.filter((node) => node.name.toLowerCase().includes(term));
});

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(1)} ${units[unitIndex]}`;
}

function formatDate(isoTimestamp: string): string {
  return new Date(isoTimestamp).toLocaleDateString();
}

function triggerFileInput(): void {
  fileInputRef.value?.click();
}

function onFileInputChange(event: Event): void {
  const input = event.target as HTMLInputElement;
  const files = input.files ? Array.from(input.files) : [];
  if (files.length > 0) store.uploadFiles(files);
  input.value = '';
}

function onDrop(event: DragEvent): void {
  isDragActive.value = false;
  const files = event.dataTransfer?.files ? Array.from(event.dataTransfer.files) : [];
  if (files.length > 0) store.uploadFiles(files);
}

function onNodeClick(node: OfficeNode): void {
  store.selectNode(node);
  if (node.kind === 'folder') {
    store.openFolder(node);
  } else if (node.doc_type === 'text') {
    // Entity navigation rule: a Doc's canonical detail view is the S147-3
    // editor, not the generic preview pane.
    router.push(`/dashboard/office/doc/${node.id}`);
  } else if (node.doc_type === 'sheet') {
    // Entity navigation rule: a Sheet's canonical detail view is the S147-4
    // grid, not the generic preview pane.
    router.push(`/dashboard/office/sheet/${node.id}`);
  } else {
    // Entity navigation rule: clicking a document opens its canonical detail
    // view — here, the preview pane (Space has no separate detail route yet).
    previewNode.value = node;
  }
}

async function onCreateDoc(name: string): Promise<void> {
  showNewDocDialog.value = false;
  const view = await officeDocApi.createDoc(name, store.currentParentId);
  router.push(`/dashboard/office/doc/${view.id}`);
}

async function onCreateSheet(name: string): Promise<void> {
  showNewSheetDialog.value = false;
  const view = await officeSheetApi.createSheet(name, store.currentParentId);
  router.push(`/dashboard/office/sheet/${view.id}`);
}

function onTreeSelect(selection: { id: string | null; name: string }): void {
  if (selection.id === null) {
    store.goToBreadcrumb(0);
    return;
  }
  store.navigateToNode({
    id: selection.id,
    name: selection.name,
    parent_id: null,
    kind: 'folder',
    trashed_at: null,
    created_at: '',
    updated_at: '',
  });
}

function onNodeMenu(node: OfficeNode, event: MouseEvent): void {
  contextMenuNode.value = node;
  contextMenuX.value = event.clientX;
  contextMenuY.value = event.clientY;
}

async function onContextAction(action: OfficeContextMenuAction): Promise<void> {
  const node = contextMenuNode.value;
  contextMenuNode.value = null;
  if (!node) return;

  switch (action) {
    case 'download':
      await downloadNode(node);
      break;
    case 'share':
      sharingNode.value = node;
      break;
    case 'rename':
      renamingNode.value = node;
      break;
    case 'move':
      movingNode.value = node;
      break;
    case 'versions':
      versionsNode.value = node;
      await store.fetchVersions(node);
      break;
    case 'trash':
      await store.trashNode(node);
      if (previewNode.value?.id === node.id) previewNode.value = null;
      break;
  }
}

/**
 * Fetch an authed download as a Blob and trigger it via a temporary anchor —
 * the download route is session-auth (Bearer header), which a bare
 * `<a href>` navigation cannot carry.
 */
async function downloadNode(node: OfficeNode): Promise<void> {
  const { blob, filename } = await officeApi.downloadDocument(node.id, node.name);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function onCreateFolder(name: string): Promise<void> {
  showNewFolderDialog.value = false;
  await store.createFolder(name);
}

async function onRenameConfirm(name: string): Promise<void> {
  const node = renamingNode.value;
  renamingNode.value = null;
  if (node) await store.renameNode(node, name);
}

async function onMoveConfirm(targetId: string | null): Promise<void> {
  const node = movingNode.value;
  movingNode.value = null;
  if (node) await store.moveNode(node, targetId);
}

async function onRestoreVersion(versionNo: number): Promise<void> {
  const node = versionsNode.value;
  if (node) await store.restoreVersion(node, versionNo);
}

onMounted(() => {
  store.fetchNodes();
  store.fetchUsage();
});
</script>

<style scoped>
.office-space-home {
  display: flex;
  height: 100%;
  min-height: 480px;
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  overflow: hidden;
  box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);
}
.office-sidebar {
  width: 220px;
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  border-right: 1px solid var(--vbwd-color-border, #e9ecef);
  padding: 16px 8px;
}
.office-sidebar-title {
  padding: 0 8px;
}
.office-sidebar-tree {
  flex: 1;
  overflow-y: auto;
}
.office-main {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
  position: relative;
}
.office-toolbar {
  padding: 16px 20px 0;
}
.office-toolbar-actions {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 12px;
}
.office-hidden-input {
  display: none;
}
.office-content {
  flex: 1;
  overflow-y: auto;
  padding: 0 20px 20px;
  position: relative;
}
.office-content.is-drag-active {
  outline: 2px dashed var(--vbwd-color-primary, #3498db);
  outline-offset: -8px;
}
.office-dropzone-overlay {
  position: absolute;
  inset: 0;
  background: rgba(52, 152, 219, 0.08);
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  color: var(--vbwd-color-primary, #3498db);
  pointer-events: none;
  z-index: 10;
}
.office-loading,
.office-error-banner,
.office-empty-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 8px;
  padding: 60px 20px;
  color: var(--vbwd-color-text-secondary, #666);
  text-align: center;
}
.office-empty-title {
  font-size: 1.05rem;
  font-weight: 600;
  color: var(--vbwd-color-text-primary, #2c3e50);
  margin: 8px 0 0;
}
.office-empty-hint {
  margin: 0 0 8px;
}
.office-node-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
}
.office-node-row {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 8px;
  border-bottom: 1px solid var(--vbwd-border-light, #eee);
  cursor: pointer;
}
.office-node-row:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-node-name {
  flex: 1;
  color: var(--vbwd-color-text-primary, #2c3e50);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.office-node-modified,
.office-node-size {
  width: 100px;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
  text-align: right;
}
.office-node-menu-btn {
  background: none;
  border: none;
  cursor: pointer;
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 1.1rem;
  padding: 4px 8px;
}
.office-node-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
  gap: 12px;
  padding-top: 12px;
}
.office-node-tile {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 6px;
  padding: 16px 8px;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 8px);
  cursor: pointer;
  position: relative;
}
.office-node-tile:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-node-tile .office-node-name {
  flex: none;
  font-size: 0.8rem;
  text-align: center;
  max-width: 100%;
}
.office-node-tile .office-node-menu-btn {
  position: absolute;
  top: 4px;
  right: 4px;
}
</style>
