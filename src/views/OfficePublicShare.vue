<template>
  <div
    class="office-public-page"
    data-testid="office-public-page"
  >
    <div class="office-public-card">
      <div
        v-if="state === 'loading'"
        class="office-public-loading"
        data-testid="office-public-loading"
      >
        {{ $t('office.publicShare.loading') }}
      </div>

      <div
        v-else-if="state === 'unavailable'"
        class="office-public-error"
        data-testid="office-public-error"
      >
        <p class="office-public-error-title">
          {{ $t('office.publicShare.unavailable') }}
        </p>
        <p class="office-public-error-hint">
          {{ $t('office.publicShare.unavailableHint') }}
        </p>
      </div>

      <div
        v-else-if="metadata"
        class="office-public-ready"
      >
        <p
          class="office-public-name"
          data-testid="office-public-name"
        >
          {{ metadata.name }}
        </p>
        <p
          v-if="permissionMessageKey"
          class="office-public-permission"
          data-testid="office-public-permission"
        >
          {{ $t(permissionMessageKey) }}
        </p>

        <form
          v-if="requiresPassword && !unlocked"
          class="office-public-password-form"
          data-testid="office-public-password-form"
          @submit.prevent="handleUnlock"
        >
          <p class="office-public-password-prompt">
            {{ $t('office.publicShare.passwordPrompt') }}
          </p>
          <input
            v-model="passwordInput"
            type="password"
            class="vbwd-input"
            :placeholder="$t('office.publicShare.passwordPlaceholder')"
            data-testid="office-public-password-input"
          >
          <button
            type="submit"
            class="vbwd-btn vbwd-btn--primary"
            data-testid="office-public-unlock-button"
            :disabled="unlocking"
          >
            {{ $t('office.publicShare.unlock') }}
          </button>
          <p
            v-if="passwordError"
            class="office-public-password-error"
            data-testid="office-public-password-error"
          >
            {{ passwordError }}
          </p>
        </form>

        <div
          v-else
          class="office-public-actions"
        >
          <button
            type="button"
            class="vbwd-btn vbwd-btn--primary"
            data-testid="office-public-download-button"
            :disabled="downloading"
            @click="handleDownload"
          >
            {{ $t('office.publicShare.download') }}
          </button>

          <div
            v-if="metadata.permission === 'edit'"
            class="office-public-replace"
          >
            <label class="office-public-replace-label">
              {{ $t('office.publicShare.replaceLabel') }}
            </label>
            <input
              ref="replaceInputRef"
              type="file"
              data-testid="office-public-replace-input"
              @change="handleReplaceFileSelected"
            >
            <p
              v-if="replaceSuccess"
              class="office-public-replace-success"
              data-testid="office-public-replace-success"
            >
              {{ $t('office.publicShare.replaceSuccess') }}
            </p>
            <p
              v-if="replaceError"
              class="office-public-replace-error"
              data-testid="office-public-replace-error"
            >
              {{ replaceError }}
            </p>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import {
  officePublicShareApi,
  OfficePublicPasswordRequiredError,
  OfficePublicShareUnavailableError,
  type OfficePublicShareMetadata,
} from '../api/officeApi';

/**
 * `/s/:token` — the public share view. Registered OUTSIDE the authenticated
 * layout (`meta.noLayout: true` in `index.ts`): no burger, no user chrome, no
 * store bootstrap that assumes a session. This is what a logged-out stranger
 * sees on the link, so every state here must be reachable with NO auth token
 * at all (S147-2 acceptance: a clean incognito window).
 */
const props = defineProps<{ token: string }>();

type PageState = 'loading' | 'unavailable' | 'ready';

const state = ref<PageState>('loading');
const metadata = ref<OfficePublicShareMetadata | null>(null);
const requiresPassword = computed(() => metadata.value?.requires_password ?? false);
const unlocked = ref(false);
const grantToken = ref<string | null>(null);

const passwordInput = ref('');
const passwordError = ref<string | null>(null);
const unlocking = ref(false);

const downloading = ref(false);

const replaceInputRef = ref<HTMLInputElement | null>(null);
const replaceSuccess = ref(false);
const replaceError = ref<string | null>(null);

const permissionMessageKey = computed(() => {
  switch (metadata.value?.permission) {
    case 'edit':
      return 'office.publicShare.permissionEdit';
    case 'comment':
      return 'office.publicShare.permissionComment';
    case 'view':
      return 'office.publicShare.permissionView';
    default:
      return null;
  }
});

async function loadMetadata(): Promise<void> {
  state.value = 'loading';
  try {
    metadata.value = await officePublicShareApi.getMetadata(props.token);
    state.value = 'ready';
  } catch (caught) {
    if (caught instanceof OfficePublicShareUnavailableError) {
      state.value = 'unavailable';
      return;
    }
    throw caught;
  }
}

async function handleUnlock(): Promise<void> {
  unlocking.value = true;
  passwordError.value = null;
  try {
    grantToken.value = await officePublicShareApi.unlock(props.token, passwordInput.value);
    unlocked.value = true;
  } catch (caught) {
    if (caught instanceof OfficePublicShareUnavailableError) {
      state.value = 'unavailable';
      return;
    }
    passwordError.value = (caught as Error).message;
  } finally {
    unlocking.value = false;
  }
}

async function handleDownload(): Promise<void> {
  downloading.value = true;
  try {
    const fallbackName = metadata.value?.name ?? 'download';
    const { blob, filename } = await officePublicShareApi.download(
      props.token,
      fallbackName,
      grantToken.value,
    );
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    URL.revokeObjectURL(url);
  } catch (caught) {
    if (caught instanceof OfficePublicPasswordRequiredError) {
      unlocked.value = false;
      grantToken.value = null;
    } else if (caught instanceof OfficePublicShareUnavailableError) {
      state.value = 'unavailable';
    }
  } finally {
    downloading.value = false;
  }
}

async function handleReplaceFileSelected(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file) return;

  replaceSuccess.value = false;
  replaceError.value = null;
  try {
    await officePublicShareApi.replaceContent(props.token, file, grantToken.value);
    replaceSuccess.value = true;
  } catch (caught) {
    if (caught instanceof OfficePublicPasswordRequiredError) {
      unlocked.value = false;
      grantToken.value = null;
    } else if (caught instanceof OfficePublicShareUnavailableError) {
      state.value = 'unavailable';
    } else {
      replaceError.value = (caught as Error).message;
    }
  }
}

onMounted(() => {
  void loadMetadata();
});
</script>

<style scoped>
.office-public-page {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--vbwd-color-page-bg, #f5f5f5);
  padding: 20px;
}
.office-public-card {
  width: 100%;
  max-width: 420px;
  background: var(--vbwd-color-surface, #fff);
  border-radius: var(--vbwd-card-radius, 8px);
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
  padding: 32px 24px;
}
.office-public-loading,
.office-public-error {
  text-align: center;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-public-error-title {
  font-weight: 600;
  color: var(--vbwd-color-text-primary, #2c3e50);
  margin: 0 0 8px;
}
.office-public-error-hint {
  margin: 0;
  font-size: 0.85rem;
}
.office-public-name {
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--vbwd-color-text-primary, #2c3e50);
  margin: 0 0 4px;
  word-break: break-word;
}
.office-public-permission {
  font-size: 0.82rem;
  color: var(--vbwd-color-text-secondary, #666);
  margin: 0 0 20px;
}
.office-public-password-form {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.office-public-password-prompt {
  margin: 0;
  font-size: 0.85rem;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-public-password-error {
  color: var(--vbwd-color-danger, #e74c3c);
  font-size: 0.82rem;
  margin: 0;
}
.office-public-actions {
  display: flex;
  flex-direction: column;
  gap: 16px;
}
.office-public-replace {
  border-top: 1px solid var(--vbwd-color-border, #e9ecef);
  padding-top: 16px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.office-public-replace-label {
  font-size: 0.8rem;
  font-weight: 600;
  color: var(--vbwd-color-text-secondary, #666);
}
.office-public-replace-success {
  color: var(--vbwd-color-success, #27ae60);
  font-size: 0.82rem;
  margin: 0;
}
.office-public-replace-error {
  color: var(--vbwd-color-danger, #e74c3c);
  font-size: 0.82rem;
  margin: 0;
}
</style>
