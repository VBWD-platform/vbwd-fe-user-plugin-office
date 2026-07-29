<template>
  <node-view-wrapper
    as="span"
    class="office-asset-image-wrapper"
    data-testid="office-doc-asset-image"
  >
    <img
      v-if="resolvedSrc"
      :src="resolvedSrc"
      :alt="(node.attrs.alt as string) || ''"
      class="office-asset-image"
    >
    <span
      v-else
      class="office-asset-image-placeholder"
      data-testid="office-doc-asset-image-placeholder"
    >{{ $t('office.doc.image.loading') }}</span>
  </node-view-wrapper>
</template>

<script setup lang="ts">
import { NodeViewWrapper, nodeViewProps } from '@tiptap/vue-3';
import { inject, onBeforeUnmount, onMounted, ref, watch, type Ref } from 'vue';

import { officeDocApi } from '../api/officeApi';
import { OFFICE_ASSET_SRC_PREFIX } from '../tiptap/OfficeAssetImage';

const props = defineProps(nodeViewProps);

// Provided by `OfficeDocEditor.vue` — the currently-open Doc's node id,
// which is what the asset-fetch route is scoped under
// (`/docs/<docNodeId>/assets/<assetNodeId>`).
const docNodeId = inject<Ref<string>>('officeDocNodeId');

// The backend's `get_asset` gate reads the DOCUMENT'S OWN SAVED content to
// decide whether an asset id is referenced (`doc_assets.py`'s ACL) — a
// just-inserted image is only in the EDITOR's in-memory state until the
// (debounced, ~2s) autosave lands, so the very first fetch after insertion
// routinely 404s. Retrying a few times covers that save-latency window
// without polling forever once an asset genuinely doesn't resolve (e.g. a
// stale/foreign id).
const ASSET_FETCH_RETRY_DELAYS_MS = [800, 1600, 2400, 3200];

const resolvedSrc = ref('');
let objectUrl: string | null = null;
let retryTimer: ReturnType<typeof setTimeout> | null = null;

function revokePreviousObjectUrl(): void {
  if (objectUrl) {
    URL.revokeObjectURL(objectUrl);
    objectUrl = null;
  }
}

function clearPendingRetry(): void {
  if (retryTimer) clearTimeout(retryTimer);
  retryTimer = null;
}

async function resolveSrc(): Promise<void> {
  clearPendingRetry();
  revokePreviousObjectUrl();
  resolvedSrc.value = '';
  const src = props.node.attrs.src as string | undefined;
  if (!src || !src.startsWith(OFFICE_ASSET_SRC_PREFIX) || !docNodeId?.value) return;

  const assetNodeId = src.slice(OFFICE_ASSET_SRC_PREFIX.length);
  await attemptFetch(assetNodeId, 0);
}

async function attemptFetch(assetNodeId: string, attempt: number): Promise<void> {
  try {
    const blob = await officeDocApi.fetchAssetBlob(docNodeId!.value, assetNodeId);
    objectUrl = URL.createObjectURL(blob);
    resolvedSrc.value = objectUrl;
  } catch {
    const delay = ASSET_FETCH_RETRY_DELAYS_MS[attempt];
    if (delay === undefined) return; // out of retries — placeholder stays visible
    retryTimer = setTimeout(() => attemptFetch(assetNodeId, attempt + 1), delay);
  }
}

onMounted(resolveSrc);
watch(() => props.node.attrs.src, resolveSrc);
onBeforeUnmount(() => {
  clearPendingRetry();
  revokePreviousObjectUrl();
});
</script>

<style scoped>
.office-asset-image-wrapper {
  display: inline-block;
}
.office-asset-image {
  max-width: 100%;
  border-radius: var(--vbwd-card-radius, 4px);
}
.office-asset-image-placeholder {
  display: inline-block;
  padding: 8px 12px;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
  background: var(--vbwd-bg-hover, #f5f6f7);
  border-radius: var(--vbwd-card-radius, 4px);
}
</style>
