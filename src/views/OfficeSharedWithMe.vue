<template>
  <div
    class="vbwd-card office-shared-with-me"
    data-testid="office-shared-with-me"
  >
    <div class="office-shared-with-me-header">
      <h2 class="vbwd-heading">
        {{ $t('office.sharedWithMe.title') }}
      </h2>
      <router-link
        to="/dashboard/office"
        class="office-shared-with-me-back"
        data-testid="office-shared-with-me-back"
      >
        {{ $t('office.sharedWithMe.back') }}
      </router-link>
    </div>

    <div
      v-if="loading"
      class="office-loading"
      data-testid="office-shared-with-me-loading"
    >
      {{ $t('office.loading') }}
    </div>

    <p
      v-else-if="entries.length === 0"
      class="office-shared-with-me-empty"
      data-testid="office-shared-with-me-empty"
    >
      {{ $t('office.sharedWithMe.empty') }}
    </p>

    <ul
      v-else
      class="office-shared-with-me-list"
      data-testid="office-shared-with-me-list"
    >
      <li
        v-for="entry in entries"
        :key="entry.id"
        class="office-shared-with-me-item"
        data-testid="office-shared-with-me-item"
      >
        <Icon :name="entry.kind === 'folder' ? 'layers' : 'document'" />
        <span
          class="office-shared-with-me-name"
          data-testid="office-shared-with-me-name"
        >{{ entry.name }}</span>
        <span
          class="office-shared-with-me-permission"
          data-testid="office-shared-with-me-permission"
        >{{ $t(`office.share.permission.${entry.share.permission}`) }}</span>
        <button
          v-if="entry.kind === 'document'"
          type="button"
          class="vbwd-btn vbwd-btn--sm vbwd-btn--primary"
          data-testid="office-shared-with-me-open"
          @click="openEntry(entry)"
        >
          {{ $t('office.sharedWithMe.openLink') }}
        </button>
      </li>
    </ul>
  </div>
</template>

<script setup lang="ts">
import { onMounted, ref } from 'vue';
import { Icon } from 'vbwd-view-component';
import { officeApi, type OfficeSharedWithMeEntry } from '../api/officeApi';

const entries = ref<OfficeSharedWithMeEntry[]>([]);
const loading = ref(false);

async function loadEntries(): Promise<void> {
  loading.value = true;
  try {
    entries.value = await officeApi.listSharedWithMe();
  } finally {
    loading.value = false;
  }
}

/**
 * Fetch an authed download as a Blob and trigger it via a temporary anchor —
 * mirrors `SpaceHome.vue`'s own vault download (the route is session-auth,
 * a bare `<a href>` navigation cannot carry the Bearer header).
 */
async function openEntry(entry: OfficeSharedWithMeEntry): Promise<void> {
  const { blob, filename } = await officeApi.downloadSharedDocument(entry.id, entry.name);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

onMounted(() => {
  void loadEntries();
});
</script>

<style scoped>
.office-shared-with-me {
  padding: 24px;
}
.office-shared-with-me-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}
.office-shared-with-me-back {
  color: var(--vbwd-color-primary, #3498db);
  font-size: 0.85rem;
}
.office-loading,
.office-shared-with-me-empty {
  color: var(--vbwd-color-text-secondary, #666);
  padding: 40px 0;
  text-align: center;
}
.office-shared-with-me-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
}
.office-shared-with-me-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 8px;
  border-bottom: 1px solid var(--vbwd-border-light, #eee);
}
.office-shared-with-me-name {
  flex: 1;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-shared-with-me-permission {
  font-size: 0.78rem;
  text-transform: capitalize;
  color: var(--vbwd-color-text-secondary, #666);
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  padding: 2px 8px;
}
</style>
