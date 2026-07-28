<template>
  <div
    class="office-doc-editor-page"
    data-testid="office-doc-editor"
  >
    <header class="office-doc-header">
      <router-link
        to="/dashboard/office"
        class="office-doc-back"
        data-testid="office-doc-back"
      >
        {{ $t('office.doc.backToFiles') }}
      </router-link>
      <h2 class="office-doc-name">
        {{ store.name }}
      </h2>
      <div class="office-doc-header-actions">
        <span
          class="office-doc-save-status"
          data-testid="office-doc-save-status"
        >{{ saveStatusLabel }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-doc-versions-btn"
          @click="openVersions"
        >
          {{ $t('office.versions.title') }}
        </button>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          data-testid="office-doc-ai-toggle-panel"
          @click="showAiSidebar = !showAiSidebar"
        >
          {{ $t('office.ai.title') }}
        </button>
      </div>
    </header>

    <div
      v-if="store.loading"
      class="office-doc-loading"
      data-testid="office-doc-loading"
    >
      {{ $t('office.loading') }}
    </div>

    <div
      v-else-if="store.loadError"
      class="office-doc-error"
      data-testid="office-doc-load-error"
    >
      {{ store.loadError }}
    </div>

    <template v-else>
      <div
        v-if="showViewOnlyBanner"
        class="office-doc-banner office-doc-banner--info"
        data-testid="office-doc-view-only-banner"
      >
        {{ $t('office.doc.viewOnlyBanner') }}
      </div>

      <div
        v-if="showLeaseBanner"
        class="office-doc-banner office-doc-banner--warning"
        data-testid="office-doc-lease-banner"
      >
        <span>{{ $t('office.doc.lease.otherEditing') }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
          data-testid="office-doc-lease-takeover"
          @click="onTakeover"
        >
          {{ $t('office.doc.lease.takeOver') }}
        </button>
      </div>

      <div
        v-if="store.conflict === 'stale_version'"
        class="office-doc-banner office-doc-banner--warning"
        data-testid="office-doc-conflict-banner"
      >
        <span>{{ $t('office.doc.conflict.staleVersion') }}</span>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
          data-testid="office-doc-reload"
          @click="onReload"
        >
          {{ $t('office.doc.reload') }}
        </button>
      </div>

      <div class="office-doc-body">
        <div class="office-doc-editor-column">
          <div
            v-if="editor"
            class="office-doc-toolbar"
            data-testid="office-doc-toolbar"
          >
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('bold') }"
              data-testid="office-doc-toolbar-bold"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleBold().run()"
            >
              B
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('italic') }"
              data-testid="office-doc-toolbar-italic"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleItalic().run()"
            >
              I
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('underline') }"
              data-testid="office-doc-toolbar-underline"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleUnderline().run()"
            >
              U
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('strike') }"
              data-testid="office-doc-toolbar-strike"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleStrike().run()"
            >
              S
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('heading', { level: 1 }) }"
              data-testid="office-doc-toolbar-h1"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleHeading({ level: 1 }).run()"
            >
              H1
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('heading', { level: 2 }) }"
              data-testid="office-doc-toolbar-h2"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleHeading({ level: 2 }).run()"
            >
              H2
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('bulletList') }"
              data-testid="office-doc-toolbar-bullet-list"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleBulletList().run()"
            >
              •
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('orderedList') }"
              data-testid="office-doc-toolbar-ordered-list"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleOrderedList().run()"
            >
              1.
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('blockquote') }"
              data-testid="office-doc-toolbar-blockquote"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleBlockquote().run()"
            >
              &quot;
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              :class="{ 'is-active': editor.isActive('codeBlock') }"
              data-testid="office-doc-toolbar-code-block"
              :disabled="!isEditable"
              @click="editor.chain().focus().toggleCodeBlock().run()"
            >
              {{ '</>' }}
            </button>
            <button
              type="button"
              class="office-doc-toolbar-btn"
              data-testid="office-doc-toolbar-link"
              :disabled="!isEditable"
              @click="onInsertLink"
            >
              🔗
            </button>
          </div>

          <EditorContent
            class="office-doc-content"
            data-testid="office-doc-content"
            :editor="editor"
          />
        </div>

        <OfficeDocAiSidebar
          v-if="showAiSidebar"
          :ai-enabled="store.aiEnabled"
          :can-toggle="store.access === 'owner'"
          :has-selection="hasSelection"
          :running="store.aiRunning"
          :error="store.aiError"
          :proposal="store.aiProposal"
          @toggle-enabled="onToggleAi"
          @run-capability="onRunCapability"
          @accept="onAcceptProposal"
          @discard="store.clearAiProposal()"
        />
      </div>
    </template>

    <OfficeVersionsDialog
      v-if="showVersionsDialog"
      :versions="versions"
      @close="showVersionsDialog = false"
      @restore="onRestoreVersion"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue';
import { useRoute } from 'vue-router';
import { Editor, EditorContent } from '@tiptap/vue-3';
import StarterKit from '@tiptap/starter-kit';
import Image from '@tiptap/extension-image';
import { officeApi, type OfficeAiCapability, type OfficeVersion } from '../api/officeApi';
import { useOfficeDocStore } from '../stores/useOfficeDocStore';
import OfficeDocAiSidebar from '../components/OfficeDocAiSidebar.vue';
import OfficeVersionsDialog from '../components/OfficeVersionsDialog.vue';

// A bounded window of surrounding text sent alongside the selection (data
// minimisation — the server also hard-caps this independently). Character
// count, not a magic UI constant: kept here as the ONE place the FE's half
// of "only the selection + a bounded window" is defined.
const AI_CONTEXT_WINDOW_CHARS = 800;

const route = useRoute();
const store = useOfficeDocStore();

// shallowRef, not ref — Editor is a mutable class instance managing its own
// internal (ProseMirror) reactivity; deep-reactive-wrapping it is both
// unnecessary and (per Tiptap's own `useEditor` composable, which returns a
// `ShallowRef`) the documented, type-correct way to hold one.
const editor = shallowRef<Editor | undefined>(undefined);
const showAiSidebar = ref(false);
const showVersionsDialog = ref(false);
const versions = ref<OfficeVersion[]>([]);
let autosaveTimer: ReturnType<typeof setTimeout> | null = null;
let aiPatchRange: { from: number; to: number } | null = null;

const nodeId = computed(() => String(route.params.id || ''));

const isEditable = computed(
  () => store.canEdit() && !(store.lease?.held && !store.lease?.is_self),
);

const showViewOnlyBanner = computed(() => !store.canEdit());
const showLeaseBanner = computed(
  () => store.canEdit() && !!store.lease?.held && !store.lease?.is_self,
);

const hasSelection = computed(() => !!editor.value && !editor.value.state.selection.empty);

const saveStatusLabel = computed(() => {
  if (store.saving) return 'office.doc.saving';
  if (store.conflict) return 'office.doc.unsavedChanges';
  if (store.dirty) return 'office.doc.unsavedChanges';
  return 'office.doc.saved';
});

function scheduleAutosave(): void {
  if (autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(triggerAutosave, 2000);
}

async function triggerAutosave(): Promise<void> {
  if (!editor.value || !isEditable.value || !store.dirty) return;
  const saved = await store.save(editor.value.getJSON());
  if (!saved && store.conflict === 'locked') {
    await store.refreshPresence();
  }
}

function buildEditor(): void {
  editor.value = new Editor({
    content: store.content || undefined,
    editable: isEditable.value,
    extensions: [
      StarterKit.configure({ link: { openOnClick: false } }),
      Image,
    ],
    onUpdate: () => {
      store.markDirty();
      scheduleAutosave();
    },
    onBlur: () => {
      if (store.dirty) triggerAutosave();
    },
  });
}

watch(isEditable, (editable) => {
  editor.value?.setEditable(editable);
});

async function onTakeover(): Promise<void> {
  await store.takeover();
}

async function onReload(): Promise<void> {
  await store.load(nodeId.value);
  if (store.content) editor.value?.commands.setContent(store.content);
}

function onInsertLink(): void {
  if (!editor.value) return;
  // eslint-disable-next-line no-alert
  const url = window.prompt('URL (https://...)');
  if (!url) return;
  editor.value.chain().focus().extendMarkRange('link').setLink({ href: url }).run();
}

function textWindow(from: number, to: number): { before: string; after: string } {
  if (!editor.value) return { before: '', after: '' };
  const doc = editor.value.state.doc;
  const before = doc.textBetween(Math.max(0, from - AI_CONTEXT_WINDOW_CHARS), from, '\n', '\n');
  const after = doc.textBetween(
    to,
    Math.min(doc.content.size, to + AI_CONTEXT_WINDOW_CHARS),
    '\n',
    '\n',
  );
  return { before, after };
}

async function onRunCapability(capability: OfficeAiCapability, targetLanguage?: string): Promise<void> {
  if (!editor.value) return;
  const { from, to } = editor.value.state.selection;
  const selectionText = editor.value.state.doc.textBetween(from, to, '\n', '\n');
  const { before, after } = textWindow(from, to);
  aiPatchRange = { from, to };
  await store.runAiCapability(capability, {
    selectionText,
    contextBefore: before,
    contextAfter: after,
    targetLanguage,
  });
}

function onAcceptProposal(): void {
  if (!editor.value || !store.aiProposal || !aiPatchRange) return;
  editor.value
    .chain()
    .focus()
    .insertContentAt(aiPatchRange, store.aiProposal.proposedText)
    .run();
  store.markDirty();
  scheduleAutosave();
  store.clearAiProposal();
}

async function onToggleAi(enabled: boolean): Promise<void> {
  if (!editor.value) return;
  await store.save(editor.value.getJSON(), { aiEnabled: enabled });
}

async function openVersions(): Promise<void> {
  versions.value = await officeApi.listVersions(nodeId.value);
  showVersionsDialog.value = true;
}

async function onRestoreVersion(versionNo: number): Promise<void> {
  await officeApi.restoreVersion(nodeId.value, versionNo);
  showVersionsDialog.value = false;
  await onReload();
}

onMounted(async () => {
  await store.load(nodeId.value);
  buildEditor();
});

onBeforeUnmount(() => {
  if (autosaveTimer) clearTimeout(autosaveTimer);
  editor.value?.destroy();
  store.releaseLease();
  store.stopPolling();
  store.reset();
});
</script>

<style scoped>
.office-doc-editor-page {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 480px;
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  overflow: hidden;
}
.office-doc-header {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
}
.office-doc-back {
  color: var(--vbwd-color-primary, #3498db);
  font-size: 0.85rem;
}
.office-doc-name {
  flex: 1;
  margin: 0;
  font-size: 1rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-doc-header-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}
.office-doc-save-status {
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-doc-loading,
.office-doc-error {
  padding: 40px;
  text-align: center;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-doc-banner {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 8px 16px;
  font-size: 0.85rem;
}
.office-doc-banner--info {
  background: rgba(52, 152, 219, 0.08);
  color: var(--vbwd-color-primary, #3498db);
}
.office-doc-banner--warning {
  background: rgba(243, 156, 18, 0.12);
  color: #b9770e;
}
.office-doc-body {
  flex: 1;
  display: flex;
  min-height: 0;
}
.office-doc-editor-column {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
}
.office-doc-toolbar {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  padding: 8px 16px;
  border-bottom: 1px solid var(--vbwd-color-border, #e9ecef);
}
.office-doc-toolbar-btn {
  min-width: 28px;
  height: 28px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-doc-toolbar-btn:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-doc-toolbar-btn.is-active {
  background: var(--vbwd-color-primary, #3498db);
  color: #fff;
}
.office-doc-toolbar-btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
.office-doc-content {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
}
</style>
