import { defineStore } from 'pinia';
import { ref } from 'vue';
import {
  officeDocApi,
  OfficeAiBudgetExceededError,
  OfficeAiForbiddenError,
  OfficeDocConflictError,
  type OfficeAiCapability,
  type OfficeDocContentModel,
  type OfficeDocLeaseState,
  type OfficeDocView,
} from '../api/officeApi';

const PRESENCE_POLL_INTERVAL_MS = 5000;
const LEASE_HEARTBEAT_INTERVAL_MS = 30000;

/** Access levels that may save content or hold the edit lease — mirrors the
 * backend's `EDIT_CAPABLE_ACCESS` (owner or an edit-permission share). */
const EDIT_CAPABLE_ACCESS = new Set(['owner', 'edit']);

export type OfficeDocConflictKind = 'stale_version' | 'locked' | null;

export const useOfficeDocStore = defineStore('officeDoc', () => {
  const nodeId = ref<string | null>(null);
  const name = ref('');
  const content = ref<OfficeDocContentModel | null>(null);
  const versionNo = ref(0);
  const access = ref<OfficeDocView['access'] | null>(null);
  const aiEnabled = ref(false);
  const lease = ref<OfficeDocLeaseState | null>(null);

  const loading = ref(false);
  const loadError = ref<string | null>(null);
  const saving = ref(false);
  const dirty = ref(false);
  const conflict = ref<OfficeDocConflictKind>(null);

  const aiRunning = ref(false);
  const aiError = ref<string | null>(null);
  const aiProposal = ref<{ capability: OfficeAiCapability; proposedText: string } | null>(null);

  let presenceTimer: ReturnType<typeof setInterval> | null = null;
  let heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  function canEdit(): boolean {
    return !!access.value && EDIT_CAPABLE_ACCESS.has(access.value);
  }

  function applyView(view: OfficeDocView): void {
    nodeId.value = view.id;
    name.value = view.name;
    content.value = view.content;
    versionNo.value = view.version_no;
    access.value = view.access;
    aiEnabled.value = view.document.ai_enabled;
    lease.value = view.lease;
    conflict.value = null;
  }

  async function load(id: string): Promise<void> {
    loading.value = true;
    loadError.value = null;
    try {
      const view = await officeDocApi.getDoc(id);
      applyView(view);
      dirty.value = false;
      if (canEdit()) await acquireLease();
      startPresencePolling();
      startLeaseHeartbeat();
    } catch (caught) {
      loadError.value = (caught as Error).message;
    } finally {
      loading.value = false;
    }
  }

  async function createDoc(name: string, parentId: string | null): Promise<OfficeDocView> {
    const view = await officeDocApi.createDoc(name, parentId);
    applyView(view);
    startPresencePolling();
    startLeaseHeartbeat();
    return view;
  }

  async function save(newContent: OfficeDocContentModel, options?: { aiEnabled?: boolean }): Promise<boolean> {
    if (!nodeId.value || !canEdit()) return false;
    saving.value = true;
    try {
      const version = await officeDocApi.saveDoc(
        nodeId.value,
        newContent,
        versionNo.value,
        options?.aiEnabled,
      );
      content.value = newContent;
      versionNo.value = version.version_no;
      if (options?.aiEnabled !== undefined) aiEnabled.value = options.aiEnabled;
      dirty.value = false;
      conflict.value = null;
      return true;
    } catch (caught) {
      if (caught instanceof OfficeDocConflictError) {
        conflict.value = caught.kind;
        if (caught.kind === 'locked' && lease.value) {
          lease.value = {
            ...lease.value,
            held: true,
            is_self: false,
            granted: false,
            holder_user_id: caught.holderUserId ?? lease.value.holder_user_id,
            expires_at: caught.expiresAt ?? lease.value.expires_at,
          };
        }
        return false;
      }
      throw caught;
    } finally {
      saving.value = false;
    }
  }

  function markDirty(): void {
    dirty.value = true;
  }

  async function acquireLease(): Promise<void> {
    if (!nodeId.value) return;
    try {
      lease.value = await officeDocApi.acquireLease(nodeId.value);
      if (lease.value.granted) conflict.value = null;
    } catch {
      // A failed heartbeat should not crash the editor — the next poll retries.
    }
  }

  async function releaseLease(): Promise<void> {
    if (!nodeId.value) return;
    try {
      await officeDocApi.releaseLease(nodeId.value);
    } catch {
      // Best-effort — the lease also self-expires.
    }
  }

  async function refreshPresence(): Promise<void> {
    if (!nodeId.value) return;
    try {
      lease.value = await officeDocApi.presence(nodeId.value);
    } catch {
      // A missed poll is not a product failure — the next tick retries.
    }
  }

  function startPresencePolling(): void {
    stopPresencePolling();
    presenceTimer = setInterval(refreshPresence, PRESENCE_POLL_INTERVAL_MS);
  }

  function stopPresencePolling(): void {
    if (presenceTimer) clearInterval(presenceTimer);
    presenceTimer = null;
  }

  function startLeaseHeartbeat(): void {
    stopLeaseHeartbeat();
    heartbeatTimer = setInterval(() => {
      if (canEdit() && lease.value?.is_self) acquireLease();
    }, LEASE_HEARTBEAT_INTERVAL_MS);
  }

  function stopLeaseHeartbeat(): void {
    if (heartbeatTimer) clearInterval(heartbeatTimer);
    heartbeatTimer = null;
  }

  function stopPolling(): void {
    stopPresencePolling();
    stopLeaseHeartbeat();
  }

  async function takeover(): Promise<boolean> {
    await acquireLease();
    return !!lease.value?.granted;
  }

  async function runAiCapability(
    capability: OfficeAiCapability,
    params: {
      selectionText?: string;
      contextBefore?: string;
      contextAfter?: string;
      targetLanguage?: string;
    },
  ): Promise<void> {
    if (!nodeId.value) return;
    aiRunning.value = true;
    aiError.value = null;
    aiProposal.value = null;
    try {
      const result = await officeDocApi.runAiCapability(nodeId.value, {
        capability,
        ...params,
      });
      aiProposal.value = { capability: result.capability, proposedText: result.proposed_text };
    } catch (caught) {
      if (caught instanceof OfficeAiForbiddenError) {
        aiError.value = 'office.ai.forbidden';
      } else if (caught instanceof OfficeAiBudgetExceededError) {
        aiError.value = 'office.ai.budgetExceeded';
      } else {
        aiError.value = (caught as Error).message;
      }
    } finally {
      aiRunning.value = false;
    }
  }

  function clearAiProposal(): void {
    aiProposal.value = null;
    aiError.value = null;
  }

  function reset(): void {
    stopPolling();
    nodeId.value = null;
    name.value = '';
    content.value = null;
    versionNo.value = 0;
    access.value = null;
    aiEnabled.value = false;
    lease.value = null;
    loading.value = false;
    loadError.value = null;
    saving.value = false;
    dirty.value = false;
    conflict.value = null;
    clearAiProposal();
  }

  return {
    nodeId,
    name,
    content,
    versionNo,
    access,
    aiEnabled,
    lease,
    loading,
    loadError,
    saving,
    dirty,
    conflict,
    aiRunning,
    aiError,
    aiProposal,
    canEdit,
    load,
    createDoc,
    save,
    markDirty,
    acquireLease,
    releaseLease,
    refreshPresence,
    startPresencePolling,
    stopPolling,
    takeover,
    runAiCapability,
    clearAiProposal,
    reset,
  };
});
