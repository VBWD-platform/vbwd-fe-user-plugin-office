<template>
  <div
    class="office-context-menu"
    data-testid="office-context-menu"
    :style="{ top: `${clampedY}px`, left: `${clampedX}px` }"
    @click.stop
  >
    <template v-if="node">
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
        data-testid="office-context-copy"
        @click="$emit('action', 'copy')"
      >
        {{ $t('office.contextMenu.copy') }}
      </button>
      <button
        type="button"
        class="office-context-menu-item"
        data-testid="office-context-cut"
        @click="$emit('action', 'cut')"
      >
        {{ $t('office.contextMenu.cut') }}
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
      <button
        type="button"
        class="office-context-menu-item office-context-menu-item--danger"
        data-testid="office-context-delete"
        @click="$emit('action', 'delete')"
      >
        {{ $t('office.contextMenu.delete') }}
      </button>
    </template>

    <button
      v-else
      type="button"
      class="office-context-menu-item"
      data-testid="office-context-paste"
      :disabled="!canPaste"
      @click="$emit('action', 'paste')"
    >
      {{ $t('office.contextMenu.paste') }}
    </button>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import type { OfficeNode } from '../api/officeApi';

export type OfficeContextMenuAction =
  | 'download'
  | 'share'
  | 'rename'
  | 'copy'
  | 'cut'
  | 'paste'
  | 'move'
  | 'versions'
  | 'trash'
  | 'delete';

/**
 * Nominal menu box, used to keep the WHOLE menu on screen rather than just its
 * top-left corner. Deliberately constants rather than a measured rect: the menu
 * is positioned on the same tick it is created, so it has no layout yet, and a
 * measure-then-reposition would make it visibly jump. Kept slightly generous —
 * over-clamping costs a few pixels, under-clamping costs the whole feature.
 */
const MENU_WIDTH = 170;
const VIEWPORT_MARGIN = 8;
//: One item's rendered height (padding + line-height) and the menu's own vertical
//: padding. The height estimate is DERIVED from how many items actually render,
//: not a fixed number: a fixed 200px was correct for six items and silently wrong
//: the moment copy/cut/delete were added — the menu under-clamped, slid off the
//: bottom of the viewport, and every click on it timed out again. Deriving it
//: means adding another item cannot reintroduce that bug.
const MENU_ITEM_HEIGHT = 34;
const MENU_VERTICAL_PADDING = 16;

const props = defineProps<{
  /** `null` for a right-click on empty space — Finder vocabulary: that menu
   * offers Paste only (the target folder is always the CURRENT folder, never
   * whichever row happens to be under the click). */
  node: OfficeNode | null;
  x: number;
  y: number;
  /** Whether the clipboard has an entry — gates the empty-space Paste item. */
  canPaste: boolean;
}>();
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
/** How many items this menu will render — mirrors the template's own v-ifs. */
const visibleItemCount = computed(() => {
  if (!props.node) return 1; // empty-space menu: Paste only
  const alwaysShown = 6; // share, rename, copy, cut, move, trash
  const documentOnly = props.node.kind === 'document' ? 2 : 0; // download, versions
  return alwaysShown + documentOnly + 1; // + delete
});

const estimatedMenuHeight = computed(
  () => visibleItemCount.value * MENU_ITEM_HEIGHT + MENU_VERTICAL_PADDING,
);

const clampedX = computed(() =>
  clamp(props.x, MENU_WIDTH, typeof window === 'undefined' ? Infinity : window.innerWidth),
);
const clampedY = computed(() =>
  clamp(
    props.y,
    estimatedMenuHeight.value,
    typeof window === 'undefined' ? Infinity : window.innerHeight,
  ),
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
.office-context-menu-item:disabled {
  color: var(--vbwd-color-text-secondary, #999);
  cursor: not-allowed;
}
.office-context-menu-item:disabled:hover {
  background: none;
}
.office-context-menu-item--danger {
  color: var(--vbwd-color-danger, #e74c3c);
}
</style>
