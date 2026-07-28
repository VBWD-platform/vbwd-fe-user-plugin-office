<template>
  <div
    class="office-dialog-overlay"
    data-testid="office-move-modal"
    @click.self="$emit('cancel')"
  >
    <div class="office-dialog">
      <h3 class="office-dialog-title">
        {{ $t('office.contextMenu.move') }}
      </h3>

      <p
        class="office-move-target"
        data-testid="office-move-target-label"
      >
        {{ targetName ?? $t('office.root') }}
      </p>

      <div
        class="office-move-tree"
        data-testid="office-move-tree"
      >
        <OfficeFolderTree
          :parent-id="null"
          :label="$t('office.root')"
          :depth="0"
          :active-id="targetId"
          @select="onSelect"
        />
      </div>

      <div class="office-dialog-actions">
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost"
          data-testid="office-move-cancel"
          @click="$emit('cancel')"
        >
          {{ $t('office.newFolder.cancel') }}
        </button>
        <button
          type="button"
          class="vbwd-btn vbwd-btn--primary"
          data-testid="office-move-confirm"
          :disabled="!hasSelected"
          @click="$emit('confirm', targetId)"
        >
          {{ $t('office.contextMenu.move') }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import OfficeFolderTree from './OfficeFolderTree.vue';

defineEmits<{ confirm: [targetId: string | null]; cancel: [] }>();

const targetId = ref<string | null>(null);
const targetName = ref<string | null>(null);
const hasSelected = ref(false);

function onSelect(selection: { id: string | null; name: string }): void {
  targetId.value = selection.id;
  targetName.value = selection.name;
  hasSelected.value = true;
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
.office-dialog-title {
  margin: 0 0 8px;
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 1.05rem;
}
.office-move-target {
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
  margin: 0 0 8px;
}
.office-move-tree {
  max-height: 220px;
  overflow-y: auto;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  margin-bottom: 16px;
}
.office-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
