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
      <input
        ref="inputRef"
        v-model="value"
        type="text"
        class="vbwd-input office-dialog-input"
        :placeholder="placeholder"
        :data-testid="`${testidPrefix}-input`"
        @keyup.enter="confirm"
        @keyup.esc="$emit('cancel')"
      >
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
          type="button"
          class="vbwd-btn vbwd-btn--primary"
          :data-testid="`${testidPrefix}-confirm`"
          :disabled="!value.trim()"
          @click="confirm"
        >
          {{ confirmLabel }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, ref } from 'vue';

const props = defineProps<{
  title: string;
  placeholder: string;
  confirmLabel: string;
  cancelLabel: string;
  testidPrefix: string;
  initialValue?: string;
}>();

const emit = defineEmits<{ confirm: [value: string]; cancel: [] }>();

const value = ref(props.initialValue ?? '');
const inputRef = ref<HTMLInputElement | null>(null);

onMounted(() => {
  inputRef.value?.focus();
});

function confirm(): void {
  const trimmed = value.value.trim();
  if (!trimmed) return;
  emit('confirm', trimmed);
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
  margin: 0 0 12px;
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 1.05rem;
}
.office-dialog-input {
  width: 100%;
  box-sizing: border-box;
  margin-bottom: 16px;
}
.office-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
