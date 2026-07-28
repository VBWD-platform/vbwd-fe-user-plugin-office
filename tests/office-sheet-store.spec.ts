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

const { FakeOfficeDocConflictError } = vi.hoisted(() => {
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
  return { FakeOfficeDocConflictError: HoistedOfficeDocConflictError };
});

vi.mock('../src/api/officeApi', () => ({
  officeSheetApi: {
    getSheet: (...args: unknown[]) => mockGetSheet(...args),
    createSheet: (...args: unknown[]) => mockCreateSheet(...args),
    saveCells: (...args: unknown[]) => mockSaveCells(...args),
    recalc: (...args: unknown[]) => mockRecalc(...args),
    exportWorkbook: (...args: unknown[]) => mockExportWorkbook(...args),
    importWorkbook: (...args: unknown[]) => mockImportWorkbook(...args),
  },
  officeDocApi: {
    acquireLease: (...args: unknown[]) => mockAcquireLease(...args),
    releaseLease: (...args: unknown[]) => mockReleaseLease(...args),
    presence: (...args: unknown[]) => mockPresence(...args),
  },
  OfficeDocConflictError: FakeOfficeDocConflictError,
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
  });
});
