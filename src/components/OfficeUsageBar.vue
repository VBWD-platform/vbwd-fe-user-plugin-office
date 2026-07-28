<template>
  <div
    class="office-usage"
    data-testid="office-usage-bar"
  >
    <div class="office-usage-track">
      <div
        class="office-usage-fill"
        :class="{ 'is-near-limit': fractionUsed > 0.9 }"
        :style="{ width: `${Math.min(fractionUsed * 100, 100)}%` }"
        data-testid="office-usage-fill"
      />
    </div>
    <span class="office-usage-label">
      {{ $t('office.usage.label') }}: {{ formatBytes(usage.bytes_used) }}
      {{ $t('office.usage.of') }} {{ formatBytes(usage.bytes_quota) }}
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import type { OfficeUsage } from '../api/officeApi';

const props = defineProps<{ usage: OfficeUsage }>();

const fractionUsed = computed(() =>
  props.usage.bytes_quota > 0 ? props.usage.bytes_used / props.usage.bytes_quota : 0,
);

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(1)} ${units[unitIndex]}`;
}
</script>

<style scoped>
.office-usage {
  display: flex;
  flex-direction: column;
  gap: 4px;
  padding: 12px;
}
.office-usage-track {
  height: 6px;
  border-radius: 3px;
  background: var(--vbwd-bg-muted, #eef0f2);
  overflow: hidden;
}
.office-usage-fill {
  height: 100%;
  background: var(--vbwd-color-primary, #3498db);
  transition: width 0.2s ease;
}
.office-usage-fill.is-near-limit {
  background: var(--vbwd-color-danger, #e74c3c);
}
.office-usage-label {
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
}
</style>
