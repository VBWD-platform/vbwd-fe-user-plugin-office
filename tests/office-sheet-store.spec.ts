import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';

const mockGetSheet = vi.fn();
const mockCreateSheet = vi.fn();
const mockSaveCells = vi.fn();
const mockRecalc = vi.fn();
const mockExportWorkbook = vi.fn();
const mockImportWorkbook = vi.fn();
const mockAcquireLease = vi.fn();
const mockReleaseLease = vi.fn();
const mockPresence = vi.fn();
const mockRunAiCapability = vi.fn();

const { FakeOfficeDocConflictError, FakeOfficeAiForbiddenError, FakeOfficeAiBudgetExceededError } = vi.hoisted(() => {
  class HoistedOfficeDocConflictError extends Error {
    kind: 'stale_version' | 'locked';
    holderUserId?: string;
    expiresAt?: string;
    constructor(body: { error: string; holder_user_id?: string; expires_at?: string }) {
      super('conflict');
      this.kind = body.error === 'locked' ? 'locked' : 'stale_version';
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
  officeSheetApi: {
    getSheet: (...args: unknown[]) => mockGetSheet(...args),
    createSheet: (...args: unknown[]) => mockCreateSheet(...args),
    saveCells: (...args: unknown[]) => mockSaveCells(...args),
    recalc: (...args: unknown[]) => mockRecalc(...args),
    exportWorkbook: (...args: unknown[]) => mockExportWorkbook(...args),
    importWorkbook: (...args: unknown[]) => mockImportWorkbook(...args),
    runAiCapability: (...args: unknown[]) => mockRunAiCapability(...args),
  },
  officeDocApi: {
    acquireLease: (...args: unknown[]) => mockAcquireLease(...args),
    releaseLease: (...args: unknown[]) => mockReleaseLease(...args),
    presence: (...args: unknown[]) => mockPresence(...args),
  },
  OfficeDocConflictError: FakeOfficeDocConflictError,
  OfficeAiForbiddenError: FakeOfficeAiForbiddenError,
  OfficeAiBudgetExceededError: FakeOfficeAiBudgetExceededError,
}));

import { changeFromRawInput, useOfficeSheetStore } from '../src/stores/useOfficeSheetStore';

function sheetView(overrides: Record<string, unknown> = {}) {
  return {
    id: 'sheet-1',
    parent_id: null,
    kind: 'document',
    name: 'Budget',
    trashed_at: null,
    created_at: '2026-07-28T00:00:00Z',
    updated_at: '2026-07-28T00:00:00Z',
    document: {
      id: 'document-1',
      node_id: 'sheet-1',
      doc_type: 'sheet',
      mime_type: 'application/json',
      size_bytes: 10,
      ai_enabled: false,
      current_version_id: 'v1',
      created_at: '2026-07-28T00:00:00Z',
      updated_at: '2026-07-28T00:00:00Z',
    },
    workbook: { sheets: [{ name: 'Sheet1', cells: {} }], active_sheet: 'Sheet1' },
    version_no: 1,
    access: 'owner',
    lease: { held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null },
    ...overrides,
  };
}

describe('changeFromRawInput', () => {
  it('an "=" prefix produces a formula change', () => {
    expect(changeFromRawInput('Sheet1', 'A1', '=SUM(B1:B3)')).toEqual({
      sheet: 'Sheet1',
      address: 'A1',
      formula: '=SUM(B1:B3)',
    });
  });

  it('an empty string produces a clear change', () => {
    expect(changeFromRawInput('Sheet1', 'A1', '')).toEqual({
      sheet: 'Sheet1',
      address: 'A1',
      clear: true,
    });
  });

  it('a numeric string produces a numeric value change', () => {
    expect(changeFromRawInput('Sheet1', 'A1', '42')).toEqual({
      sheet: 'Sheet1',
      address: 'A1',
      value: 42,
    });
  });

  it('a non-numeric string produces a text value change', () => {
    expect(changeFromRawInput('Sheet1', 'A1', 'hello')).toEqual({
      sheet: 'Sheet1',
      address: 'A1',
      value: 'hello',
    });
  });
});

describe('useOfficeSheetStore', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    vi.useFakeTimers();
    mockAcquireLease.mockResolvedValue({
      held: true,
      holder_user_id: 'me',
      is_self: true,
      granted: true,
      expires_at: '2026-07-28T00:01:30Z',
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

  it('load() populates the workbook and acquires a lease for an edit-capable access', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    const store = useOfficeSheetStore();

    await store.load('sheet-1');

    expect(store.workbook).toEqual(sheetView().workbook);
    expect(store.versionNo).toBe(1);
    expect(store.activeSheetName).toBe('Sheet1');
    expect(mockAcquireLease).toHaveBeenCalledWith('sheet-1');
    store.stopPolling();
  });

  it('load() does NOT attempt a lease for a view-only access', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ access: 'view' }));
    const store = useOfficeSheetStore();

    await store.load('sheet-1');

    expect(mockAcquireLease).not.toHaveBeenCalled();
    expect(store.canEdit()).toBe(false);
    store.stopPolling();
  });

  it('setCellFromRawInput echoes the value locally, then flushes a save that returns only the deltas', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!A1': 5 } });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setCellFromRawInput('Sheet1', 'A1', '5');
    expect(store.getCell('Sheet1', 'A1').v).toBe(5);
    expect(store.dirty).toBe(true);

    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', address: 'A1', value: 5 }],
      1,
    );
    expect(store.versionNo).toBe(2);
    store.stopPolling();
  });

  it('setCellFromRawInput with a formula recalculates a dependent cell via the server response', async () => {
    mockGetSheet.mockResolvedValue(
      sheetView({
        workbook: {
          sheets: [{ name: 'Sheet1', cells: { A1: { v: 2 }, B1: { v: 3 } } }],
          active_sheet: 'Sheet1',
        },
      }),
    );
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!C1': 5 } });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setCellFromRawInput('Sheet1', 'C1', '=A1+B1');
    await vi.advanceTimersByTimeAsync(500);

    expect(store.getCell('Sheet1', 'C1').v).toBe(5);
    expect(store.getCell('Sheet1', 'C1').f).toBe('=A1+B1');
    store.stopPolling();
  });

  it('a stale_version conflict is surfaced without throwing', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockRejectedValue(new FakeOfficeDocConflictError({ error: 'stale_version' }));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setCellFromRawInput('Sheet1', 'A1', '5');
    await vi.advanceTimersByTimeAsync(500);

    expect(store.conflict).toBe('stale_version');
    store.stopPolling();
  });

  it('a locked conflict updates the lease view', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockRejectedValue(
      new FakeOfficeDocConflictError({
        error: 'locked',
        holder_user_id: 'someone-else',
        expires_at: '2026-07-28T00:05:00Z',
      }),
    );
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setCellFromRawInput('Sheet1', 'A1', '5');
    await vi.advanceTimersByTimeAsync(500);

    expect(store.conflict).toBe('locked');
    expect(store.lease?.holder_user_id).toBe('someone-else');
    store.stopPolling();
  });

  it('a view-only store never queues a cell change', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ access: 'view' }));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setCellFromRawInput('Sheet1', 'A1', '5');
    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).not.toHaveBeenCalled();
    store.stopPolling();
  });

  it('addSheetTab adds an empty tab locally and switches to it', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.addSheetTab('Sheet2');

    expect(store.getSheetTab('Sheet2')).toBeDefined();
    expect(store.activeSheetName).toBe('Sheet2');
    await vi.advanceTimersByTimeAsync(500);
    store.stopPolling();
  });

  it('recalcAll() applies the server deltas', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockRecalc.mockResolvedValue({ version_no: 3, changes: { 'Sheet1!A1': 9 } });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.recalcAll();

    expect(store.versionNo).toBe(3);
    expect(store.getCell('Sheet1', 'A1').v).toBe(9);
    store.stopPolling();
  });

  it('reset() clears every field back to its initial state', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.reset();

    expect(store.nodeId).toBeNull();
    expect(store.workbook).toBeNull();
    expect(store.access).toBeNull();
    expect(store.aiEnabled).toBe(false);
    expect(store.aiProposal).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// S147-3.5 — the Sheet AI helper.
// ---------------------------------------------------------------------------

describe('useOfficeSheetStore — AI helper', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    vi.useFakeTimers();
    mockAcquireLease.mockResolvedValue({
      held: true, holder_user_id: 'me', is_self: true, granted: true,
      expires_at: '2026-07-30T00:01:30Z',
    });
    mockPresence.mockResolvedValue({
      held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null,
    });
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('applyView reads ai_enabled from the document into the store', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');
    expect(store.aiEnabled).toBe(true);
    store.stopPolling();
  });

  it('toggleAi saves an empty-changes request carrying ai_enabled and updates local state', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.toggleAi(true);

    expect(mockSaveCells).toHaveBeenCalledWith('sheet-1', [], 1, true);
    expect(store.aiEnabled).toBe(true);
    expect(store.versionNo).toBe(2);
    store.stopPolling();
  });

  it('toggleAi does nothing for a caller who cannot edit', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ access: 'view' }));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.toggleAi(true);

    expect(mockSaveCells).not.toHaveBeenCalled();
    expect(store.aiEnabled).toBe(false);
    store.stopPolling();
  });

  it('runAiCapability populates aiProposal on success', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'formula',
      capability: 'sheet_write_formula',
      connection_slug: 'default',
      address: 'B1',
      formula: '=SUM(A1:A2)',
    });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.runAiCapability('sheet_write_formula', { address: 'B1', intent: 'sum column A' });

    expect(mockRunAiCapability).toHaveBeenCalledWith('sheet-1', {
      capability: 'sheet_write_formula',
      address: 'B1',
      intent: 'sum column A',
    });
    expect(store.aiProposal?.formula).toBe('=SUM(A1:A2)');
    expect(store.aiRunning).toBe(false);
    store.stopPolling();
  });

  it('runAiCapability forwards a prompt through to officeSheetApi for sheet_freeform', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'formula',
      capability: 'sheet_freeform',
      connection_slug: 'default',
      address: 'B1',
      formula: '=B1*2',
    });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.runAiCapability('sheet_freeform', {
      address: 'B1',
      prompt: 'add a column that is column B times 2',
    });

    expect(mockRunAiCapability).toHaveBeenCalledWith('sheet-1', {
      capability: 'sheet_freeform',
      address: 'B1',
      prompt: 'add a column that is column B times 2',
    });
    expect(store.aiProposal?.formula).toBe('=B1*2');
    store.stopPolling();
  });

  it('runAiCapability maps a forbidden error onto a translatable key', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockRejectedValue(new FakeOfficeAiForbiddenError('nope'));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.runAiCapability('sheet_explain_formula', { address: 'B1' });

    expect(store.aiError).toBe('office.sheet.ai.forbidden');
    expect(store.aiProposal).toBeNull();
    store.stopPolling();
  });

  it('runAiCapability maps a budget-exceeded error onto a translatable key', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockRejectedValue(new FakeOfficeAiBudgetExceededError('nope'));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    await store.runAiCapability('sheet_explain_formula', { address: 'B1' });

    expect(store.aiError).toBe('office.sheet.ai.budgetExceeded');
    store.stopPolling();
  });

  it('acceptAiProposal for a formula proposal applies it through saveCells, never a raw write', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'formula',
      capability: 'sheet_write_formula',
      connection_slug: 'default',
      address: 'B1',
      formula: '=SUM(A1:A2)',
    });
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!B1': 30 } });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');
    await store.runAiCapability('sheet_write_formula', { address: 'B1', intent: 'sum' });

    const applied = await store.acceptAiProposal();

    expect(applied).toBe(true);
    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', address: 'B1', formula: '=SUM(A1:A2)' }],
      1,
    );
    expect(store.getCell('Sheet1', 'B1').v).toBe(30);
    expect(store.aiProposal).toBeNull();
    store.stopPolling();
  });

  it('acceptAiProposal always sends the formula slot even without a leading "="', async () => {
    // Proves the model's reply can never fall through to the bare-value
    // slot `changeFromRawInput` would pick for text with no leading "=" —
    // it must go through the engine's formula parser either way.
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'formula',
      capability: 'sheet_write_formula',
      connection_slug: 'default',
      address: 'B1',
      formula: 'SUM(A1:A2)',
    });
    mockSaveCells.mockResolvedValue({
      version_no: 2,
      changes: { 'Sheet1!B1': { t: 'error', v: '#NAME?' } },
    });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');
    await store.runAiCapability('sheet_write_formula', { address: 'B1', intent: 'sum' });

    await store.acceptAiProposal();

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', address: 'B1', formula: '=SUM(A1:A2)' }],
      1,
    );
    store.stopPolling();
  });

  it('acceptAiProposal for a text proposal just clears it — nothing to apply', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'text',
      capability: 'sheet_explain_formula',
      connection_slug: 'default',
      text: 'It sums A1 and A2.',
    });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');
    await store.runAiCapability('sheet_explain_formula', { address: 'B1' });

    const applied = await store.acceptAiProposal();

    expect(applied).toBe(true);
    expect(mockSaveCells).not.toHaveBeenCalled();
    expect(store.aiProposal).toBeNull();
    store.stopPolling();
  });
});
