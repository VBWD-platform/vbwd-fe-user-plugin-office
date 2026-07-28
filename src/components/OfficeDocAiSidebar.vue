<template>
  <aside
    class="office-doc-ai-sidebar"
    data-testid="office-doc-ai-sidebar"
  >
    <div class="office-doc-ai-header">
      <h3 class="office-doc-ai-title">
        {{ $t('office.ai.title') }}
      </h3>
      <label class="office-doc-ai-toggle">
        <input
          type="checkbox"
          data-testid="office-doc-ai-enabled-toggle"
          :checked="aiEnabled"
          :disabled="!canToggle"
          @change="$emit('toggle-enabled', ($event.target as HTMLInputElement).checked)"
        >
        {{ $t('office.ai.enableToggle') }}
      </label>
    </div>

    <p
      v-if="!aiEnabled"
      class="office-doc-ai-hint"
      data-testid="office-doc-ai-disabled-hint"
    >
      {{ $t('office.ai.disabledHint') }}
    </p>

    <template v-else>
      <p class="office-doc-ai-hint">
        {{ $t('office.ai.selectTextHint') }}
      </p>

      <div class="office-doc-ai-capabilities">
        <button
          v-for="capability in capabilities"
          :key="capability.id"
          type="button"
          class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
          :data-testid="`office-doc-ai-capability-${capability.id}`"
          :disabled="running || (capability.requiresSelection && !hasSelection)"
          @click="$emit('run-capability', capability.id)"
        >
          {{ $t(`office.ai.capability.${capability.id}`) }}
        </button>
      </div>

      <div
        v-if="running"
        class="office-doc-ai-running"
        data-testid="office-doc-ai-running"
      >
        {{ $t('office.ai.running') }}
      </div>

      <div
        v-else-if="error"
        class="office-doc-ai-error"
        data-testid="office-doc-ai-error"
      >
        {{ $t(error) }}
      </div>

      <div
        v-else-if="proposal"
        class="office-doc-ai-proposal"
        data-testid="office-doc-ai-proposal"
      >
        <p class="office-doc-ai-proposal-text">
          {{ proposal.proposedText }}
        </p>
        <div class="office-doc-ai-proposal-actions">
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary vbwd-btn--sm"
            data-testid="office-doc-ai-accept"
            @click="$emit('accept')"
          >
            {{ $t('office.ai.accept') }}
          </button>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-doc-ai-discard"
            @click="$emit('discard')"
          >
            {{ $t('office.ai.discard') }}
          </button>
        </div>
      </div>
    </template>
  </aside>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import type { OfficeAiCapability } from '../api/officeApi';

const CAPABILITIES_REQUIRING_SELECTION: OfficeAiCapability[] = [
  'rewrite_shorter',
  'rewrite_longer',
  'rewrite_formal',
  'rewrite_plain',
  'summarize',
  'fix_grammar',
  'translate',
];

const ALL_CAPABILITIES: OfficeAiCapability[] = [
  'continue_writing',
  'rewrite_shorter',
  'rewrite_longer',
  'rewrite_formal',
  'rewrite_plain',
  'summarize',
  'fix_grammar',
  'translate',
  'outline',
];

defineProps<{
  aiEnabled: boolean;
  canToggle: boolean;
  hasSelection: boolean;
  running: boolean;
  error: string | null;
  proposal: { capability: OfficeAiCapability; proposedText: string } | null;
}>();

defineEmits<{
  'toggle-enabled': [enabled: boolean];
  'run-capability': [capability: OfficeAiCapability];
  accept: [];
  discard: [];
}>();

const capabilities = computed(() =>
  ALL_CAPABILITIES.map((id) => ({
    id,
    requiresSelection: CAPABILITIES_REQUIRING_SELECTION.includes(id),
  })),
);
</script>

<style scoped>
.office-doc-ai-sidebar {
  width: 280px;
  flex-shrink: 0;
  border-left: 1px solid var(--vbwd-color-border, #e9ecef);
  padding: 16px;
  overflow-y: auto;
}
.office-doc-ai-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 12px;
}
.office-doc-ai-title {
  margin: 0;
  font-size: 1rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-doc-ai-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-doc-ai-hint {
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-doc-ai-capabilities {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 12px;
}
.office-doc-ai-running,
.office-doc-ai-error {
  font-size: 0.85rem;
  padding: 8px;
  border-radius: var(--vbwd-card-radius, 6px);
}
.office-doc-ai-error {
  color: var(--vbwd-color-danger, #e74c3c);
  background: rgba(231, 76, 60, 0.08);
}
.office-doc-ai-proposal {
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 6px);
  padding: 10px;
}
.office-doc-ai-proposal-text {
  white-space: pre-wrap;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  margin: 0 0 10px;
}
.office-doc-ai-proposal-actions {
  display: flex;
  gap: 8px;
}
</style>
