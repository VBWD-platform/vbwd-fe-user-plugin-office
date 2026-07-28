<template>
  <nav
    class="office-breadcrumb"
    data-testid="office-breadcrumb"
    :aria-label="$t('office.title')"
  >
    <template
      v-for="(crumb, index) in crumbs"
      :key="crumb.id ?? 'root'"
    >
      <button
        type="button"
        class="office-breadcrumb-item"
        :class="{ 'is-current': index === crumbs.length - 1 }"
        :disabled="index === crumbs.length - 1"
        data-testid="office-breadcrumb-item"
        @click="$emit('navigate', index)"
      >
        {{ index === 0 ? $t('office.root') : crumb.name }}
      </button>
      <span
        v-if="index < crumbs.length - 1"
        class="office-breadcrumb-sep"
        aria-hidden="true"
      >/</span>
    </template>
  </nav>
</template>

<script setup lang="ts">
import type { OfficeBreadcrumbEntry } from '../stores/useOfficeStore';

defineProps<{ crumbs: OfficeBreadcrumbEntry[] }>();
defineEmits<{ navigate: [index: number] }>();
</script>

<style scoped>
.office-breadcrumb {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 4px;
  margin-bottom: 12px;
}
.office-breadcrumb-item {
  background: none;
  border: none;
  padding: 2px 4px;
  font-size: 0.9rem;
  color: var(--vbwd-color-primary, #3498db);
  cursor: pointer;
  border-radius: var(--vbwd-btn-radius, 4px);
}
.office-breadcrumb-item:hover:not(:disabled) {
  background: var(--vbwd-bg-hover, #f5f6f7);
}
.office-breadcrumb-item.is-current {
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-weight: 600;
  cursor: default;
}
.office-breadcrumb-sep {
  color: var(--vbwd-color-text-secondary, #9ca3af);
}
</style>
