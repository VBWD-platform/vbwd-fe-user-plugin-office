<template>
  <aside
    class="office-sheet-ai-sidebar"
    data-testid="office-sheet-ai-sidebar"
  >
    <div class="office-sheet-ai-header">
      <h3 class="office-sheet-ai-title">
        {{ $t('office.sheet.ai.title') }}
      </h3>
      <label class="office-sheet-ai-toggle">
        <input
          type="checkbox"
          data-testid="office-sheet-ai-enabled-toggle"
          :checked="aiEnabled"
          :disabled="!canToggle"
          @change="$emit('toggle-enabled', ($event.target as HTMLInputElement).checked)"
        >
        {{ $t('office.sheet.ai.enableToggle') }}
      </label>
    </div>

    <p
      v-if="!aiEnabled"
      class="office-sheet-ai-hint"
      data-testid="office-sheet-ai-disabled-hint"
    >
      {{ $t('office.sheet.ai.disabledHint') }}
    </p>

    <template v-else>
      <p
        v-if="!hasActiveCell"
        class="office-sheet-ai-hint"
        data-testid="office-sheet-ai-select-cell-hint"
      >
        {{ $t('office.sheet.ai.selectCellHint') }}
      </p>

      <template v-else>
        <label class="office-sheet-ai-intent-label">
          {{ $t('office.sheet.ai.intentLabel') }}
          <input
            v-model="intentText"
            type="text"
            class="vbwd-input office-sheet-ai-intent-input"
            data-testid="office-sheet-ai-intent-input"
            :placeholder="$t('office.sheet.ai.intentPlaceholder')"
          >
        </label>

        <div class="office-sheet-ai-capabilities">
          <button
            v-for="capability in capabilities"
            :key="capability.id"
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            :data-testid="`office-sheet-ai-capability-${capability.id}`"
            :disabled="running || capability.disabled"
            @click="$emit('run-capability', capability.id, intentText)"
          >
            {{ $t(`office.sheet.ai.capability.${capability.id}`) }}
          </button>
        </div>

        <div class="office-sheet-ai-freeform">
          <label
            class="office-sheet-ai-prompt-label"
            for="office-sheet-ai-prompt-input"
          >
            {{ $t('office.sheet.ai.promptLabel') }}
          </label>
          <textarea
            id="office-sheet-ai-prompt-input"
            v-model="promptText"
            class="vbwd-input office-sheet-ai-prompt-input"
            data-testid="office-sheet-ai-prompt-input"
            rows="3"
            :placeholder="$t('office.sheet.ai.promptPlaceholder')"
          />
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary vbwd-btn--sm"
            data-testid="office-sheet-ai-prompt-run"
            :disabled="running || promptText.trim() === ''"
            @click="onRunFreeform"
          >
            {{ $t('office.sheet.ai.promptRun') }}
          </button>
        </div>
      </template>

      <div
        v-if="running"
        class="office-sheet-ai-running"
        data-testid="office-sheet-ai-running"
      >
        {{ $t('office.sheet.ai.running') }}
      </div>

      <div
        v-else-if="error"
        class="office-sheet-ai-error"
        data-testid="office-sheet-ai-error"
      >
        {{ $t(error) }}
      </div>

      <div
        v-else-if="proposal"
        class="office-sheet-ai-proposal"
        data-testid="office-sheet-ai-proposal"
      >
        <p
          v-if="proposal.kind === 'formula'"
          class="office-sheet-ai-proposal-text"
          data-testid="office-sheet-ai-proposal-formula"
        >
          {{ $t('office.sheet.ai.proposedFormulaLabel') }} {{ proposal.address }}:
          <code>{{ proposal.formula }}</code>
        </p>
        <p
          v-else
          class="office-sheet-ai-proposal-text"
          data-testid="office-sheet-ai-proposal-text"
        >
          {{ proposal.text }}
        </p>
        <div class="office-sheet-ai-proposal-actions">
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary vbwd-btn--sm"
            data-testid="office-sheet-ai-accept"
            @click="$emit('accept')"
          >
            {{ $t('office.sheet.ai.accept') }}
          </button>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--ghost vbwd-btn--sm"
            data-testid="office-sheet-ai-discard"
            @click="$emit('discard')"
          >
            {{ $t('office.sheet.ai.discard') }}
          </button>
        </div>
      </div>
    </template>
  </aside>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import type { OfficeSheetAiCapability, OfficeSheetAiProposal } from '../api/officeApi';

//: The two capabilities whose reply is a FORMULA — they need edit access
//: (proposing something to APPLY) and, for `sheet_write_formula`, a
//: non-empty intent (the server rejects an empty one too; disabling the
//: button here is just the friendlier first line of defence).
const FORMULA_PRODUCING_CAPABILITIES: OfficeSheetAiCapability[] = [
  'sheet_write_formula',
  'sheet_fix_error',
];

const ALL_CAPABILITIES: OfficeSheetAiCapability[] = [
  'sheet_write_formula',
  'sheet_fix_error',
  'sheet_explain_formula',
  'sheet_summarize_range',
];

const props = defineProps<{
  aiEnabled: boolean;
  canToggle: boolean;
  canProposeFormula: boolean;
  hasActiveCell: boolean;
  running: boolean;
  error: string | null;
  proposal: OfficeSheetAiProposal | null;
}>();

const emit = defineEmits<{
  'toggle-enabled': [enabled: boolean];
  'run-capability': [capability: OfficeSheetAiCapability, intent: string];
  'run-freeform': [prompt: string];
  accept: [];
  discard: [];
}>();

const intentText = ref('');
const promptText = ref('');

function onRunFreeform(): void {
  const trimmedPrompt = promptText.value.trim();
  if (!trimmedPrompt) return;
  emit('run-freeform', trimmedPrompt);
}

const capabilities = computed(() =>
  ALL_CAPABILITIES.map((id) => {
    const isFormulaProducing = FORMULA_PRODUCING_CAPABILITIES.includes(id);
    const needsIntent = id === 'sheet_write_formula';
    const disabled =
      (isFormulaProducing && !props.canProposeFormula) ||
      (needsIntent && intentText.value.trim() === '');
    return { id, disabled };
  }),
);
</script>

<style scoped>
.office-sheet-ai-sidebar {
  width: 280px;
  flex-shrink: 0;
  border-left: 1px solid var(--vbwd-color-border, #e9ecef);
  padding: 16px;
  overflow-y: auto;
}
.office-sheet-ai-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 12px;
}
.office-sheet-ai-title {
  margin: 0;
  font-size: 1rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-sheet-ai-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-sheet-ai-hint {
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-sheet-ai-intent-label {
  display: block;
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
  margin-bottom: 10px;
}
.office-sheet-ai-intent-input {
  width: 100%;
  margin-top: 4px;
}
.office-sheet-ai-capabilities {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 12px;
}
.office-sheet-ai-freeform {
  margin-bottom: 12px;
}
.office-sheet-ai-prompt-label {
  display: block;
  font-size: 0.75rem;
  color: var(--vbwd-color-text-secondary, #666);
  margin-bottom: 4px;
}
.office-sheet-ai-prompt-input {
  width: 100%;
  resize: vertical;
  margin-bottom: 6px;
}
.office-sheet-ai-running,
.office-sheet-ai-error {
  font-size: 0.85rem;
  padding: 8px;
  border-radius: var(--vbwd-card-radius, 6px);
}
.office-sheet-ai-error {
  color: var(--vbwd-color-danger, #e74c3c);
  background: rgba(231, 76, 60, 0.08);
}
.office-sheet-ai-proposal {
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-card-radius, 6px);
  padding: 10px;
}
.office-sheet-ai-proposal-text {
  white-space: pre-wrap;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
  margin: 0 0 10px;
}
.office-sheet-ai-proposal-actions {
  display: flex;
  gap: 8px;
}
</style>
