import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';

const mockGetDoc = vi.fn();
const mockCreateDoc = vi.fn();
const mockSaveDoc = vi.fn();
const mockAcquireLease = vi.fn();
const mockReleaseLease = vi.fn();
const mockPresence = vi.fn();
const mockRunAiCapability = vi.fn();

const { FakeOfficeDocConflictError, FakeOfficeAiForbiddenError, FakeOfficeAiBudgetExceededError } =
  vi.hoisted(() => {
    class HoistedOfficeDocConflictError extends Error {
      kind: 'stale_version' | 'locked';
      currentVersionNo?: number;
      holderUserId?: string;
      expiresAt?: string;
      constructor(body: {
        error: string;
        current_version_no?: number;
        holder_user_id?: string;
        expires_at?: string;
      }) {
        super('conflict');
        this.kind = body.error === 'locked' ? 'locked' : 'stale_version';
        this.currentVersionNo = body.current_version_no;
        this.holderUserId = body.holder_user_id;
        this.expiresAt = body.expires_at;
      }
    }
    class HoistedOfficeAiForbiddenError extends Error {}
    class HoistedOfficeAiBudgetExceededError extends Error {}
    return {
      FakeOfficeDocConflictError: HoistedOfficeDocConflictError,
      FakeOfficeAiForbiddenError: HoistedOfficeAiForbiddenError,
      FakeOfficeAiBudgetExceededError: HoistedOfficeAiBudgetExceededError,
    };
  });

vi.mock('../src/api/officeApi', () => ({
  officeDocApi: {
    getDoc: (...args: unknown[]) => mockGetDoc(...args),
    createDoc: (...args: unknown[]) => mockCreateDoc(...args),
    saveDoc: (...args: unknown[]) => mockSaveDoc(...args),
    acquireLease: (...args: unknown[]) => mockAcquireLease(...args),
    releaseLease: (...args: unknown[]) => mockReleaseLease(...args),
    presence: (...args: unknown[]) => mockPresence(...args),
    runAiCapability: (...args: unknown[]) => mockRunAiCapability(...args),
  },
  OfficeDocConflictError: FakeOfficeDocConflictError,
  OfficeAiForbiddenError: FakeOfficeAiForbiddenError,
  OfficeAiBudgetExceededError: FakeOfficeAiBudgetExceededError,
}));

import { useOfficeDocStore } from '../src/stores/useOfficeDocStore';

const EMPTY_CONTENT = { type: 'doc', content: [{ type: 'paragraph' }] };

function docView(overrides: Record<string, unknown> = {}) {
  return {
    id: 'doc-1',
    parent_id: null,
    kind: 'document',
    name: 'Untitled',
    trashed_at: null,
    created_at: '2026-07-27T00:00:00Z',
    updated_at: '2026-07-27T00:00:00Z',
    document: {
      id: 'document-1',
      node_id: 'doc-1',
      doc_type: 'text',
      mime_type: 'text/plain',
      size_bytes: 10,
      ai_enabled: false,
      current_version_id: 'v1',
      created_at: '2026-07-27T00:00:00Z',
      updated_at: '2026-07-27T00:00:00Z',
    },
    content: EMPTY_CONTENT,
    version_no: 1,
    access: 'owner',
    lease: { held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null },
    ...overrides,
  };
}

describe('useOfficeDocStore', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    mockAcquireLease.mockResolvedValue({
      held: true,
      holder_user_id: 'me',
      is_self: true,
      granted: true,
      expires_at: '2026-07-27T00:01:30Z',
    });
    mockPresence.mockResolvedValue({
      held: false,
      holder_user_id: null,
      is_self: false,
      granted: true,
      expires_at: null,
    });
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('load() populates the document and acquires a lease for an edit-capable access', async () => {
    mockGetDoc.mockResolvedValue(docView());
    const store = useOfficeDocStore();

    await store.load('doc-1');

    expect(store.content).toEqual(EMPTY_CONTENT);
    expect(store.versionNo).toBe(1);
    expect(store.access).toBe('owner');
    expect(mockAcquireLease).toHaveBeenCalledWith('doc-1');
    store.stopPolling();
  });

  it('load() does NOT attempt a lease for a view-only access', async () => {
    mockGetDoc.mockResolvedValue(docView({ access: 'view' }));
    const store = useOfficeDocStore();

    await store.load('doc-1');

    expect(mockAcquireLease).not.toHaveBeenCalled();
    expect(store.canEdit()).toBe(false);
    store.stopPolling();
  });

  it('save() updates versionNo and clears dirty on success', async () => {
    mockGetDoc.mockResolvedValue(docView());
    mockSaveDoc.mockResolvedValue({ id: 'v2', version_no: 2 });
    const store = useOfficeDocStore();
    await store.load('doc-1');
    store.markDirty();

    const result = await store.save(EMPTY_CONTENT);

    expect(result).toBe(true);
    expect(store.versionNo).toBe(2);
    expect(store.dirty).toBe(false);
    store.stopPolling();
  });

  it('save() surfaces a stale_version conflict without throwing', async () => {
    mockGetDoc.mockResolvedValue(docView());
    mockSaveDoc.mockRejectedValue(
      new FakeOfficeDocConflictError({ error: 'stale_version', current_version_no: 5 }),
    );
    const store = useOfficeDocStore();
    await store.load('doc-1');

    const result = await store.save(EMPTY_CONTENT);

    expect(result).toBe(false);
    expect(store.conflict).toBe('stale_version');
    store.stopPolling();
  });

  it('save() surfaces a locked conflict and updates the lease view', async () => {
    mockGetDoc.mockResolvedValue(docView());
    mockSaveDoc.mockRejectedValue(
      new FakeOfficeDocConflictError({
        error: 'locked',
        holder_user_id: 'someone-else',
        expires_at: '2026-07-27T00:05:00Z',
      }),
    );
    const store = useOfficeDocStore();
    await store.load('doc-1');

    const result = await store.save(EMPTY_CONTENT);

    expect(result).toBe(false);
    expect(store.conflict).toBe('locked');
    expect(store.lease?.holder_user_id).toBe('someone-else');
    expect(store.lease?.is_self).toBe(false);
    store.stopPolling();
  });

  it('a non-edit-capable store never calls save at all', async () => {
    mockGetDoc.mockResolvedValue(docView({ access: 'view' }));
    const store = useOfficeDocStore();
    await store.load('doc-1');

    const result = await store.save(EMPTY_CONTENT);

    expect(result).toBe(false);
    expect(mockSaveDoc).not.toHaveBeenCalled();
    store.stopPolling();
  });

  it('takeover() re-acquires the lease and reports whether it was granted', async () => {
    mockGetDoc.mockResolvedValue(
      docView({
        lease: { held: true, holder_user_id: 'someone-else', is_self: false, granted: false, expires_at: null },
      }),
    );
    const store = useOfficeDocStore();
    await store.load('doc-1');

    const granted = await store.takeover();

    expect(granted).toBe(true);
    expect(store.lease?.is_self).toBe(true);
    store.stopPolling();
  });

  it('runAiCapability stores a proposal on success', async () => {
    mockGetDoc.mockResolvedValue(docView({ document: { ...docView().document, ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      capability: 'summarize',
      connection_slug: 'default',
      proposed_text: 'a short summary',
    });
    const store = useOfficeDocStore();
    await store.load('doc-1');

    await store.runAiCapability('summarize', { selectionText: 'hello world' });

    expect(store.aiProposal).toEqual({ capability: 'summarize', proposedText: 'a short summary' });
    expect(store.aiError).toBeNull();
    store.stopPolling();
  });

  it('runAiCapability forwards a prompt through to officeDocApi for freeform', async () => {
    mockGetDoc.mockResolvedValue(docView({ document: { ...docView().document, ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      capability: 'freeform',
      connection_slug: 'default',
      proposed_text: '- point one\n- point two',
    });
    const store = useOfficeDocStore();
    await store.load('doc-1');

    await store.runAiCapability('freeform', {
      selectionText: 'The quick brown fox',
      prompt: 'rewrite this as bullet points',
    });

    expect(mockRunAiCapability).toHaveBeenCalledWith(
      'doc-1',
      expect.objectContaining({ capability: 'freeform', prompt: 'rewrite this as bullet points' }),
    );
    expect(store.aiProposal).toEqual({ capability: 'freeform', proposedText: '- point one\n- point two' });
    store.stopPolling();
  });

  it('runAiCapability maps a forbidden error to a translation key, not a raw message', async () => {
    mockGetDoc.mockResolvedValue(docView());
    mockRunAiCapability.mockRejectedValue(new FakeOfficeAiForbiddenError('nope'));
    const store = useOfficeDocStore();
    await store.load('doc-1');

    await store.runAiCapability('summarize', { selectionText: 'hi' });

    expect(store.aiProposal).toBeNull();
    expect(store.aiError).toBe('office.ai.forbidden');
    store.stopPolling();
  });

  it('runAiCapability maps a budget-exceeded error to a translation key', async () => {
    mockGetDoc.mockResolvedValue(docView());
    mockRunAiCapability.mockRejectedValue(new FakeOfficeAiBudgetExceededError('nope'));
    const store = useOfficeDocStore();
    await store.load('doc-1');

    await store.runAiCapability('summarize', { selectionText: 'hi' });

    expect(store.aiError).toBe('office.ai.budgetExceeded');
    store.stopPolling();
  });

  it('reset() clears every field back to its initial state', async () => {
    mockGetDoc.mockResolvedValue(docView());
    const store = useOfficeDocStore();
    await store.load('doc-1');

    store.reset();

    expect(store.nodeId).toBeNull();
    expect(store.content).toBeNull();
    expect(store.access).toBeNull();
  });
});
