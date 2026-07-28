<template>
  <div
    v-if="uploads.length > 0"
    class="office-upload-progress-list"
    data-testid="office-upload-progress-list"
  >
    <div
      v-for="item in uploads"
      :key="item.id"
      class="office-upload-progress-item"
      :class="`is-${item.status}`"
      data-testid="office-upload-progress-item"
    >
      <div class="office-upload-progress-row">
        <span class="office-upload-progress-name">{{ item.name }}</span>
        <button
          v-if="item.status !== 'uploading'"
          type="button"
          class="office-upload-progress-dismiss"
          data-testid="office-upload-progress-dismiss"
          @click="$emit('dismiss', item.id)"
        >
          ×
        </button>
      </div>
      <div
        v-if="item.status === 'uploading'"
        class="office-upload-progress-track"
      >
        <div
          class="office-upload-progress-fill"
          data-testid="office-upload-progress-bar"
          :style="{ width: `${item.progress * 100}%` }"
        />
      </div>
      <span
        v-if="item.status === 'error'"
        class="office-upload-progress-error"
        data-testid="office-upload-progress-error"
      >
        {{ item.errorMessage ?? $t('office.upload.failed', { name: item.name }) }}
      </span>
    </div>
  </div>
</template>

<script setup lang="ts">
import type { OfficeUploadItem } from '../stores/useOfficeStore';

defineProps<{ uploads: OfficeUploadItem[] }>();
defineEmits<{ dismiss: [uploadId: string] }>();
</script>

<style scoped>
.office-upload-progress-list {
  position: fixed;
  right: 20px;
  bottom: 20px;
  width: 280px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  z-index: 60;
}
.office-upload-progress-item {
  background: var(--vbwd-color-surface, #fff);
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 8px);
  padding: 10px 12px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}
.office-upload-progress-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}
.office-upload-progress-name {
  font-size: 0.8rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.office-upload-progress-dismiss {
  background: none;
  border: none;
  cursor: pointer;
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 1rem;
  line-height: 1;
}
.office-upload-progress-track {
  margin-top: 6px;
  height: 4px;
  border-radius: 2px;
  background: var(--vbwd-bg-muted, #eef0f2);
  overflow: hidden;
}
.office-upload-progress-fill {
  height: 100%;
  background: var(--vbwd-color-primary, #3498db);
  transition: width 0.15s linear;
}
.office-upload-progress-item.is-error .office-upload-progress-name {
  color: var(--vbwd-color-danger, #e74c3c);
}
.office-upload-progress-error {
  display: block;
  margin-top: 4px;
  font-size: 0.75rem;
  color: var(--vbwd-color-danger, #e74c3c);
}
</style>
