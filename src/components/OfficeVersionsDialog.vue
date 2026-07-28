<template>
  <div
    class="office-dialog-overlay"
    data-testid="office-versions-modal"
    @click.self="$emit('close')"
  >
    <div class="office-dialog office-dialog--wide">
      <h3 class="office-dialog-title">
        {{ $t('office.versions.title') }}
      </h3>

      <p
        v-if="versions.length === 0"
        class="office-versions-empty"
        data-testid="office-versions-empty"
      >
        {{ $t('office.versions.empty') }}
      </p>
      <ul
        v-else
        class="office-versions-list"
        data-testid="office-versions-list"
      >
        <li
          v-for="version in versionsNewestFirst"
          :key="version.id"
          class="office-versions-item"
          data-testid="office-versions-item"
        >
          <span>v{{ version.version_no }} — {{ formatDate(version.created_at) }}</span>
          <span
            v-if="version.version_no === currentVersionNo"
            class="office-versions-current"
          >{{ $t('office.versions.current') }}</span>
          <button
            v-else
            type="button"
            class="vbwd-btn vbwd-btn--sm vbwd-btn--ghost"
            data-testid="office-versions-restore"
            @click="$emit('restore', version.version_no)"
          >
            {{ $t('office.versions.restore') }}
          </button>
        </li>
      </ul>

      <div class="office-dialog-actions">
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost"
          data-testid="office-versions-close"
          @click="$emit('close')"
        >
          {{ $t('office.versions.close') }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import type { OfficeVersion } from '../api/officeApi';

const props = defineProps<{ versions: OfficeVersion[] }>();
defineEmits<{ close: []; restore: [versionNo: number] }>();

// The backend returns versions in ASCENDING order
// (office_version_repository.py: `.order_by(OfficeVersion.version_no.asc())`)
// — this dialog displays newest-first (the usual "version history" reading
// order), so the prop is re-sorted for display rather than assumed to
// already be in the order the UI wants.
const versionsNewestFirst = computed(() =>
  [...props.versions].sort((a, b) => b.version_no - a.version_no),
);

// The CURRENT version is the highest version_no — i.e. the first entry once
// sorted newest-first, computed independently of array position so a future
// display-order change can't silently mislabel it again.
const currentVersionNo = computed(() =>
  props.versions.reduce((max, version) => Math.max(max, version.version_no), 0),
);

function formatDate(isoTimestamp: string): string {
  return new Date(isoTimestamp).toLocaleString();
}
</script>

<style scoped>
.office-dialog-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 100;
}
.office-dialog {
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  padding: 20px;
  width: 100%;
  max-width: 360px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
}
.office-dialog--wide {
  max-width: 460px;
}
.office-dialog-title {
  margin: 0 0 12px;
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 1.05rem;
}
.office-versions-empty {
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 0.85rem;
}
.office-versions-list {
  list-style: none;
  margin: 0 0 16px;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-height: 260px;
  overflow-y: auto;
}
.office-versions-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 10px;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  font-size: 0.85rem;
}
.office-versions-current {
  font-size: 0.75rem;
  color: var(--vbwd-color-success, #27ae60);
  font-weight: 600;
}
.office-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
