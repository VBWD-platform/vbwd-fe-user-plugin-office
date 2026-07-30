<template>
  <div
    class="office-dialog-overlay"
    :data-testid="`${testidPrefix}-modal`"
    @click.self="$emit('cancel')"
  >
    <div class="office-dialog">
      <h3 class="office-dialog-title">
        {{ title }}
      </h3>
      <p
        class="office-dialog-message"
        :data-testid="`${testidPrefix}-message`"
      >
        {{ message }}
      </p>

      <div class="office-dialog-actions">
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost"
          :data-testid="`${testidPrefix}-cancel`"
          @click="$emit('cancel')"
        >
          {{ cancelLabel }}
        </button>
        <button
          ref="confirmRef"
          type="button"
          class="vbwd-btn"
          :class="danger ? 'vbwd-btn--danger' : 'vbwd-btn--primary'"
          :data-testid="`${testidPrefix}-confirm`"
          @click="$emit('confirm')"
        >
          {{ confirmLabel }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, ref } from 'vue';

/**
 * A yes/no dialog for a destructive action — the counterpart to
 * `OfficeInputDialog`, which collects a value.
 *
 * It exists because a permanent delete must be confirmed and the plugin uses no
 * `window.confirm`: a native confirm cannot be styled, cannot be translated, and
 * cannot be driven by a test. Deliberately minimal — a title, a message and two
 * buttons — rather than a general-purpose modal abstraction nothing else needs.
 */
withDefaults(
  defineProps<{
    title: string;
    message: string;
    confirmLabel: string;
    cancelLabel: string;
    testidPrefix: string;
    /** Styles confirm as destructive. Default true: this dialog's whole reason
     *  for existing is destructive actions. */
    danger?: boolean;
  }>(),
  { danger: true },
);
defineEmits<{ confirm: []; cancel: [] }>();

const confirmRef = ref<HTMLButtonElement | null>(null);

onMounted(() => {
  // Focus Confirm so the dialog is keyboard-operable, but note it does NOT
  // auto-submit on Enter elsewhere — a destructive action should take a
  // deliberate keystroke on a focused button, not a stray Enter.
  confirmRef.value?.focus();
});
</script>

<style scoped>
.office-dialog-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.35);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
}
.office-dialog {
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  padding: 20px;
  min-width: 320px;
  max-width: 440px;
  box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
}
.office-dialog-title {
  margin: 0 0 8px;
  font-size: 1rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-dialog-message {
  margin: 0 0 16px;
  font-size: 0.875rem;
  line-height: 1.5;
  color: var(--vbwd-color-text-secondary, #6b7684);
}
.office-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
