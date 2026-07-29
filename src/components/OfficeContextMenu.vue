<template>
  <div
    class="office-context-menu"
    data-testid="office-context-menu"
    :style="{ top: `${clampedY}px`, left: `${clampedX}px` }"
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
import { computed } from 'vue';
import type { OfficeNode } from '../api/officeApi';

export type OfficeContextMenuAction = 'download' | 'share' | 'rename' | 'move' | 'versions' | 'trash';

/**
 * Nominal menu box, used to keep the WHOLE menu on screen rather than just its
 * top-left corner. Deliberately constants rather than a measured rect: the menu
 * is positioned on the same tick it is created, so it has no layout yet, and a
 * measure-then-reposition would make it visibly jump. Kept slightly generous —
 * over-clamping costs a few pixels, under-clamping costs the whole feature.
 */
const MENU_WIDTH = 170;
const MENU_HEIGHT = 200;
const VIEWPORT_MARGIN = 8;

const props = defineProps<{ node: OfficeNode; x: number; y: number }>();
defineEmits<{ action: [action: OfficeContextMenuAction] }>();

function clamp(requested: number, size: number, available: number): number {
  // Never negative: on a viewport smaller than the menu, flush to the edge and
  // let the menu scroll rather than positioning it off the top/left.
  return Math.max(VIEWPORT_MARGIN, Math.min(requested, available - size - VIEWPORT_MARGIN));
}

// The bug this guards: `position: fixed` at the raw pointer put the menu at
// y=723 in a 720px viewport for a row near the bottom of the list. It rendered,
// it was "visible", and it was completely unclickable — elementFromPoint
// returned null and every click timed out. Silent and total.
const clampedX = computed(() =>
  clamp(props.x, MENU_WIDTH, typeof window === 'undefined' ? Infinity : window.innerWidth),
);
const clampedY = computed(() =>
  clamp(props.y, MENU_HEIGHT, typeof window === 'undefined' ? Infinity : window.innerHeight),
);
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
