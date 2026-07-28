<template>
  <div class="office-folder-tree-node">
    <div
      class="office-folder-tree-row"
      :class="{ 'is-active': isActive }"
      :style="{ paddingLeft: `${depth * 14 + 8}px` }"
      data-testid="office-tree-row"
    >
      <button
        type="button"
        class="office-folder-tree-toggle"
        data-testid="office-tree-toggle"
        :aria-label="expanded ? 'Collapse' : 'Expand'"
        @click="toggleExpanded"
      >
        {{ expanded ? '▾' : '▸' }}
      </button>
      <button
        type="button"
        class="office-folder-tree-label"
        data-testid="office-tree-label"
        @click="$emit('select', { id: parentId, name: label })"
      >
        {{ label }}
      </button>
    </div>

    <div
      v-if="expanded"
      class="office-folder-tree-children"
    >
      <p
        v-if="loading"
        class="office-folder-tree-loading"
      >
        {{ $t('office.loading') }}
      </p>
      <OfficeFolderTree
        v-for="folder in childFolders"
        :key="folder.id"
        :parent-id="folder.id"
        :label="folder.name"
        :depth="depth + 1"
        :active-id="activeId"
        @select="$emit('select', $event)"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import { officeApi, type OfficeNode } from '../api/officeApi';

const props = defineProps<{
  parentId: string | null;
  label: string;
  depth: number;
  activeId: string | null;
}>();

defineEmits<{ select: [{ id: string | null; name: string }] }>();

const expanded = ref(false);
const loading = ref(false);
const childFolders = ref<OfficeNode[]>([]);

const isActive = computed(() => props.activeId === props.parentId);

async function toggleExpanded(): Promise<void> {
  expanded.value = !expanded.value;
  if (expanded.value && childFolders.value.length === 0) {
    loading.value = true;
    try {
      const nodes = await officeApi.listNodes(props.parentId);
      childFolders.value = nodes.filter((node) => node.kind === 'folder');
    } finally {
      loading.value = false;
    }
  }
}
</script>

<style scoped>
.office-folder-tree-row {
  display: flex;
  align-items: center;
  gap: 4px;
  border-radius: var(--vbwd-btn-radius, 4px);
}
.office-folder-tree-row.is-active {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-folder-tree-toggle {
  background: none;
  border: none;
  width: 16px;
  cursor: pointer;
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 0.7rem;
}
.office-folder-tree-label {
  background: none;
  border: none;
  text-align: left;
  flex: 1;
  padding: 6px 4px;
  cursor: pointer;
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 0.85rem;
}
.office-folder-tree-label:hover {
  color: var(--vbwd-color-primary, #3498db);
}
.office-folder-tree-loading {
  padding-left: 28px;
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
  margin: 4px 0;
}
</style>
