<template>
  <div
    class="office-dialog-overlay"
    data-testid="office-share-dialog"
    @click.self="$emit('close')"
  >
    <div class="office-dialog office-dialog--wide">
      <h3 class="office-dialog-title">
        {{ $t('office.share.title', { name: node.name }) }}
      </h3>

      <form
        class="office-share-form"
        data-testid="office-share-form"
        @submit.prevent="handleCreate"
      >
        <label class="office-share-field-label">
          {{ $t('office.share.recipientLabel') }}
        </label>
        <input
          v-model="recipientEmail"
          type="email"
          class="vbwd-input"
          :placeholder="$t('office.share.recipientPlaceholder')"
          data-testid="office-share-recipient-input"
        >

        <fieldset class="office-share-permission-group">
          <legend class="office-share-field-label">
            {{ $t('office.share.permissionLabel') }}
          </legend>
          <label
            v-for="option in permissionOptions"
            :key="option"
            class="office-share-permission-option"
          >
            <input
              v-model="permission"
              type="radio"
              name="office-share-permission"
              :value="option"
              :data-testid="`office-share-permission-${option}`"
            >
            {{ $t(`office.share.permission.${option}`) }}
          </label>
        </fieldset>

        <label class="office-share-anonymous-row">
          <input
            v-model="allowAnonymous"
            type="checkbox"
            data-testid="office-share-anonymous-toggle"
          >
          <span>
            <strong>{{ $t('office.share.anonymousLabel') }}</strong>
            <br>
            <span class="office-share-anonymous-hint">{{ $t('office.share.anonymousHint') }}</span>
          </span>
        </label>

        <div class="office-share-row">
          <label class="office-share-field-label">
            {{ $t('office.share.expiryLabel') }}
          </label>
          <input
            v-model="expiresAt"
            type="datetime-local"
            class="vbwd-input"
            data-testid="office-share-expiry-input"
          >
        </div>

        <div class="office-share-row">
          <label class="office-share-field-label">
            {{ $t('office.share.passwordLabel') }}
          </label>
          <input
            v-model="password"
            type="password"
            class="vbwd-input"
            :placeholder="$t('office.share.passwordPlaceholder')"
            data-testid="office-share-password-input"
          >
        </div>

        <p
          v-if="createError"
          class="office-share-error"
          data-testid="office-share-error"
        >
          {{ createError }}
        </p>

        <button
          type="submit"
          class="vbwd-btn vbwd-btn--primary"
          data-testid="office-share-create-button"
          :disabled="creating"
        >
          {{ $t('office.share.create') }}
        </button>
      </form>

      <div
        v-if="createdShare"
        class="office-share-token-panel"
        data-testid="office-share-token-display"
      >
        <p class="office-share-token-note">
          {{ $t('office.share.tokenShownOnce') }}
        </p>
        <div class="office-share-link-row">
          <input
            :value="createdLink"
            type="text"
            readonly
            class="vbwd-input"
            data-testid="office-share-link-input"
            @click="($event.target as HTMLInputElement).select()"
          >
          <button
            type="button"
            class="vbwd-btn vbwd-btn--sm vbwd-btn--ghost"
            data-testid="office-share-copy-button"
            @click="copyLink"
          >
            {{ copied ? $t('office.share.copied') : $t('office.share.copy') }}
          </button>
        </div>
      </div>

      <h4 class="office-share-list-title">
        {{ $t('office.share.existingTitle') }}
      </h4>
      <p
        v-if="shares.length === 0"
        class="office-share-empty"
        data-testid="office-share-empty"
      >
        {{ $t('office.share.empty') }}
      </p>
      <ul
        v-else
        class="office-share-list"
        data-testid="office-share-list"
      >
        <li
          v-for="share in shares"
          :key="share.id"
          class="office-share-item"
          data-testid="office-share-item"
        >
          <span class="office-share-item-permission">{{ $t(`office.share.permission.${share.permission}`) }}</span>
          <span class="office-share-item-detail">
            {{ share.subject_user_id ? $t('office.share.namedRecipient') : $t('office.share.linkShare') }}
            <template v-if="share.allow_anonymous"> · {{ $t('office.share.anonymousBadge') }}</template>
            <template v-if="share.has_password"> · {{ $t('office.share.passwordBadge') }}</template>
          </span>
          <button
            type="button"
            class="vbwd-btn vbwd-btn--sm vbwd-btn--ghost office-share-revoke-button"
            data-testid="office-share-revoke-button"
            @click="handleRevoke(share.id)"
          >
            {{ $t('office.share.revoke') }}
          </button>
        </li>
      </ul>

      <div class="office-dialog-actions">
        <button
          type="button"
          class="vbwd-btn vbwd-btn--ghost"
          data-testid="office-share-close-button"
          @click="$emit('close')"
        >
          {{ $t('office.share.close') }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import {
  officeApi,
  type OfficeNode,
  type OfficeShare,
  type OfficeSharePermission,
} from '../api/officeApi';

const props = defineProps<{ node: OfficeNode }>();
defineEmits<{ close: [] }>();

const permissionOptions: OfficeSharePermission[] = ['view', 'comment', 'edit'];

const shares = ref<OfficeShare[]>([]);
const recipientEmail = ref('');
const permission = ref<OfficeSharePermission>('view');
const allowAnonymous = ref(false);
const expiresAt = ref('');
const password = ref('');
const creating = ref(false);
const createError = ref<string | null>(null);
const createdShare = ref<{ id: string; token: string } | null>(null);
const copied = ref(false);

const createdLink = computed(() =>
  createdShare.value ? `${window.location.origin}/s/${createdShare.value.token}` : '',
);

async function loadShares(): Promise<void> {
  shares.value = await officeApi.listShares(props.node.id);
}

async function handleCreate(): Promise<void> {
  creating.value = true;
  createError.value = null;
  try {
    const created = await officeApi.createShare(props.node.id, {
      permission: permission.value,
      recipient_email: recipientEmail.value.trim() || undefined,
      allow_anonymous: allowAnonymous.value,
      password: password.value || undefined,
      // A bare `datetime-local` value has no timezone; treat it as local time
      // and let `Date` normalise to an ISO string the backend parses.
      expires_at: expiresAt.value ? new Date(expiresAt.value).toISOString() : null,
    });
    createdShare.value = { id: created.id, token: created.token };
    copied.value = false;
    recipientEmail.value = '';
    password.value = '';
    await loadShares();
  } catch (caught) {
    createError.value = (caught as Error).message;
  } finally {
    creating.value = false;
  }
}

async function handleRevoke(shareId: string): Promise<void> {
  await officeApi.revokeShare(shareId);
  if (createdShare.value?.id === shareId) createdShare.value = null;
  await loadShares();
}

async function copyLink(): Promise<void> {
  try {
    await navigator.clipboard.writeText(createdLink.value);
    copied.value = true;
  } catch {
    // Clipboard access can be denied by the browser; the link is still
    // selectable/readable in the input field, so this is not fatal.
    copied.value = false;
  }
}

onMounted(() => {
  void loadShares();
});
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
  max-height: 90vh;
  overflow-y: auto;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
}
.office-dialog--wide {
  max-width: 480px;
}
.office-dialog-title {
  margin: 0 0 12px;
  color: var(--vbwd-color-text-primary, #2c3e50);
  font-size: 1.05rem;
}
.office-share-form {
  display: flex;
  flex-direction: column;
  gap: 10px;
  margin-bottom: 16px;
}
.office-share-field-label {
  font-size: 0.8rem;
  font-weight: 600;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-share-permission-group {
  border: none;
  padding: 0;
  margin: 0;
  display: flex;
  gap: 14px;
}
.office-share-permission-option {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-share-anonymous-row {
  display: flex;
  gap: 10px;
  align-items: flex-start;
  padding: 10px;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  font-size: 0.85rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-share-anonymous-hint {
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 0.78rem;
}
.office-share-row {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.office-share-error {
  color: var(--vbwd-color-danger, #e74c3c);
  font-size: 0.82rem;
  margin: 0;
}
.office-share-token-panel {
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  padding: 10px;
  margin-bottom: 16px;
}
.office-share-token-note {
  margin: 0 0 8px;
  font-size: 0.8rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-share-link-row {
  display: flex;
  gap: 8px;
}
.office-share-link-row .vbwd-input {
  flex: 1;
}
.office-share-list-title {
  margin: 0 0 8px;
  font-size: 0.9rem;
  color: var(--vbwd-color-text-primary, #2c3e50);
}
.office-share-empty {
  color: var(--vbwd-color-text-secondary, #666);
  font-size: 0.85rem;
  margin: 0 0 16px;
}
.office-share-list {
  list-style: none;
  margin: 0 0 16px;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-height: 200px;
  overflow-y: auto;
}
.office-share-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  border: 1px solid var(--vbwd-color-border, #e9ecef);
  border-radius: var(--vbwd-btn-radius, 4px);
  font-size: 0.82rem;
}
.office-share-item-permission {
  font-weight: 600;
  color: var(--vbwd-color-text-primary, #2c3e50);
  text-transform: capitalize;
}
.office-share-item-detail {
  flex: 1;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
</style>
