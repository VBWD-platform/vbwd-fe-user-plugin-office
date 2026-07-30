<template>
  <aside
    class="office-preview"
    data-testid="office-preview-pane"
  >
    <div class="office-preview-header">
      <span class="office-preview-name">{{ node.name }}</span>
      <button
        type="button"
        class="office-preview-close"
        data-testid="office-preview-close"
        :aria-label="$t('versions.close')"
        @click="$emit('close')"
      >
        ×
      </button>
    </div>

    <div
      v-if="loading"
      class="office-preview-loading"
      data-testid="office-preview-loading"
    >
      {{ $t('office.loading') }}
    </div>
    <div
      v-else-if="loadFailed"
      class="office-preview-unavailable"
      data-testid="office-preview-error"
    >
      <p>{{ $t('office.preview.loadFailed') }}</p>
    </div>

    <img
      v-else-if="isImage"
      class="office-preview-image"
      data-testid="office-preview-image"
      :src="objectUrl"
      :alt="node.name"
    >
    <iframe
      v-else-if="isPdf"
      class="office-preview-pdf"
      data-testid="office-preview-pdf"
      :src="objectUrl"
      :title="node.name"
    />
    <div
      v-else-if="isText"
      class="office-preview-text"
      data-testid="office-preview-text"
    >
      {{ textContent }}
    </div>
    <div
      v-else
      class="office-preview-unavailable"
      data-testid="office-preview-unavailable"
    >
      <p>{{ $t('office.preview.unavailable') }}</p>
      <p>{{ $t('office.preview.downloadInstead') }}</p>
    </div>
  </aside>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, ref, watch } from 'vue';
import type { OfficeNode } from '../api/officeApi';
import { officeApi } from '../api/officeApi';

/** Server-side preview allow-list (B4 in S147-1) — anything else is download-only. */
const PREVIEWABLE_IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
const PREVIEWABLE_PDF_TYPE = 'application/pdf';
const PREVIEWABLE_TEXT_TYPE = 'text/plain';

const props = defineProps<{ node: OfficeNode }>();
defineEmits<{ close: [] }>();

const isImage = computed(() => PREVIEWABLE_IMAGE_TYPES.includes(props.node.mime_type ?? ''));
const isPdf = computed(() => props.node.mime_type === PREVIEWABLE_PDF_TYPE);
const isText = computed(() => props.node.mime_type === PREVIEWABLE_TEXT_TYPE);

const textContent = ref('');
const objectUrl = ref('');
const loading = ref(false);
const loadFailed = ref(false);

/**
 * Images and PDFs are fetched WITH the auth header and shown from an object URL.
 *
 * They used to be a bare `:src="officeApi.contentUrl(id)"`, which cannot work: a
 * browser does not attach the Authorization header to an `<img>`/`<iframe>`
 * request, so the content route answered 401 and the pane rendered an empty box
 * — the pane appeared to flash and vanish. The text branch below was always
 * correct because it fetched the blob itself; this brings the other two into
 * line, the same way the Docs editor resolves `office-asset:` images.
 */
function revokeObjectUrl(): void {
  if (objectUrl.value) {
    URL.revokeObjectURL(objectUrl.value);
    objectUrl.value = '';
  }
}

watch(
  () => props.node.id,
  async () => {
    textContent.value = '';
    revokeObjectUrl();
    loadFailed.value = false;

    if (!isImage.value && !isPdf.value && !isText.value) return;

    loading.value = true;
    try {
      const { blob } = await officeApi.downloadDocument(props.node.id, props.node.name);
      if (isText.value) {
        textContent.value = await blob.text();
      } else {
        objectUrl.value = URL.createObjectURL(blob);
      }
    } catch {
      // Say so rather than showing an empty pane — a blank preview is
      // indistinguishable from a broken app.
      loadFailed.value = true;
    } finally {
      loading.value = false;
    }
  },
  { immediate: true },
);

// Object URLs pin the blob in memory until revoked; clicking through a folder of
// images would otherwise leak one per preview.
onBeforeUnmount(revokeObjectUrl);
</script>

<style scoped>
.office-preview {
  display: flex;
  flex-direction: column;
  width: 320px;
  flex-shrink: 0;
  background: var(--vbwd-color-surface, #fff);
  border-left: 1px solid var(--vbwd-color-border, #e9ecef);
  padding: 16px;
}
.office-preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 12px;
}
.office-preview-name {
  font-weight: 600;
  color: var(--vbwd-color-text-primary, #2c3e50);
  word-break: break-all;
}
.office-preview-close {
  background: none;
  border: none;
  font-size: 1.2rem;
  line-height: 1;
  cursor: pointer;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-preview-image {
  max-width: 100%;
  border-radius: var(--vbwd-btn-radius, 4px);
}
.office-preview-pdf {
  width: 100%;
  height: 400px;
  border: none;
}
.office-preview-text {
  white-space: pre-wrap;
  font-family: monospace;
  font-size: 0.8rem;
  max-height: 400px;
  overflow: auto;
}
.office-preview-unavailable {
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 0.85rem;
}
</style>
