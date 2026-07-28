<template>
  <div
    class="office-context-menu"
    data-testid="office-context-menu"
    :style="{ top: `${y}px`, left: `${x}px` }"
    @click.stop
  >
    <button
      v-if="node.kind === 'document'"
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-download"
      @click="$emit('action', 'download')"
    >
      {{ $t('office.contextMenu.download') }}
    </button>
    <button
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-share"
      @click="$emit('action', 'share')"
    >
      {{ $t('office.contextMenu.share') }}
    </button>
    <button
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-rename"
      @click="$emit('action', 'rename')"
    >
      {{ $t('office.contextMenu.rename') }}
    </button>
    <button
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-move"
      @click="$emit('action', 'move')"
    >
      {{ $t('office.contextMenu.move') }}
    </button>
    <button
      v-if="node.kind === 'document'"
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-versions"
      @click="$emit('action', 'versions')"
    >
      {{ $t('office.contextMenu.versions') }}
    </button>
    <button
      type="button"
      class="office-context-menu-item office-context-menu-item--danger"
      data-testid="office-context-trash"
      @click="$emit('action', 'trash')"
    >
      {{ $t('office.contextMenu.trash') }}
    </button>
  </div>
</template>

<script setup lang="ts">
import type { OfficeNode } from '../api/officeApi';

export type OfficeContextMenuAction = 'download' | 'share' | 'rename' | 'move' | 'versions' | 'trash';

defineProps<{ node: OfficeNode; x: number; y: number }>();
defineEmits<{ action: [action: OfficeContextMenuAction] }>();
</script>

<style scoped>
.office-context-menu {
  position: fixed;
  z-index: 50;
  display: flex;
  flex-direction: column;
  min-width: 160px;
  background: var(--vbwd-color-surface, #fff);
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 8px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
  padding: 4px;
}
.office-context-menu-item {
  background: none;
  border: none;
  text-align: left;
  padding: 8px 10px;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  border-radius: var(--vbwd-btn-radius, 4px);
  cursor: pointer;
}
.office-context-menu-item:hover {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-context-menu-item--danger {
  color: var(--vbwd-color-danger, #e74c3c);
}
</style>
