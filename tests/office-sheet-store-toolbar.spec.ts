import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';

const mockGetSheet = vi.fn();
const mockSaveCells = vi.fn();
const mockAcquireLease = vi.fn();
const mockReleaseLease = vi.fn();
const mockPresence = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeSheetApi: {
    getSheet: (...args: unknown[]) => mockGetSheet(...args),
    createSheet: vi.fn(),
    saveCells: (...args: unknown[]) => mockSaveCells(...args),
    recalc: vi.fn(),
    exportWorkbook: vi.fn(),
    importWorkbook: vi.fn(),
  },
  officeDocApi: {
    acquireLease: (...args: unknown[]) => mockAcquireLease(...args),
    releaseLease: (...args: unknown[]) => mockReleaseLease(...args),
    presence: (...args: unknown[]) => mockPresence(...args),
  },
  OfficeDocConflictError: class extends Error {},
}));

import { useOfficeSheetStore } from '../src/stores/useOfficeSheetStore';

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
    workbook: {
      sheets: [{ name: 'Sheet1', cells: { A1: { v: 1234.5 }, B1: { v: 1 } } }],
      active_sheet: 'Sheet1',
    },
    version_no: 1,
    access: 'owner',
    lease: { held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null },
    ...overrides,
  };
}

describe('useOfficeSheetStore — toolbar / fill-handle presentation edits', () => {
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

  it('toggleBold echoes bold locally and queues a style change', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.toggleBold('Sheet1', ['A1']);

    expect(store.getCellStyle('Sheet1', 'A1').bold).toBe(true);
    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', address: 'A1', style: { bold: true } }],
      1,
    );
    store.stopPolling();
  });

  it('toggleBold a second time removes bold (a pure toggle)', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.toggleBold('Sheet1', ['A1']);
    await vi.advanceTimersByTimeAsync(500);
    store.toggleBold('Sheet1', ['A1']);
    await vi.advanceTimersByTimeAsync(500);

    expect(store.getCellStyle('Sheet1', 'A1').bold).toBe(false);
    store.stopPolling();
  });

  it('setNumberFormat("currency") seeds a default decimals count', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setNumberFormat('Sheet1', ['A1'], 'currency');
    await vi.advanceTimersByTimeAsync(500);

    expect(store.getCellStyle('Sheet1', 'A1')).toEqual({ format: 'currency', decimals: 2 });
    store.stopPolling();
  });

  it('adjustDecimals increases and clamps at the backend ceiling of 10', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.setNumberFormat('Sheet1', ['A1'], 'number');
    await vi.advanceTimersByTimeAsync(500);
    for (let tick = 0; tick < 20; tick += 1) {
      store.adjustDecimals('Sheet1', ['A1'], 1);
    }
    await vi.advanceTimersByTimeAsync(500);

    expect(store.getCellStyle('Sheet1', 'A1').decimals).toBe(10);
    store.stopPolling();
  });

  it('mergeRange queues a merge change and re-fetches the workbook afterward', async () => {
    mockGetSheet.mockResolvedValueOnce(sheetView()).mockResolvedValueOnce(
      sheetView({
        workbook: {
          sheets: [{ name: 'Sheet1', cells: {}, merges: ['A1:C1'] }],
          active_sheet: 'Sheet1',
        },
        version_no: 2,
      }),
    );
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.mergeRange('Sheet1', 'A1:C1');
    await vi.advanceTimersByTimeAsync(500);
    await vi.waitFor(() => expect(mockGetSheet).toHaveBeenCalledTimes(2));

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', merge: 'A1:C1' }],
      1,
    );
    expect(store.getMerges('Sheet1')).toEqual(['A1:C1']);
    store.stopPolling();
  });

  it('unmergeAt queues an unmerge change', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.unmergeAt('Sheet1', 'B1');
    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', unmerge: 'B1' }],
      1,
    );
    store.stopPolling();
  });

  it('fillCells queues one fill_from change per target address', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!A2': 1234.5, 'Sheet1!A3': 1234.5 } });
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.fillCells('Sheet1', 'A1', ['A2', 'A3']);
    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [
        { sheet: 'Sheet1', address: 'A2', fill_from: 'A1' },
        { sheet: 'Sheet1', address: 'A3', fill_from: 'A1' },
      ],
      1,
    );
    expect(store.getCell('Sheet1', 'A2').v).toBe(1234.5);
    store.stopPolling();
  });

  it('a view-only store never queues a presentation change', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ access: 'view' }));
    const store = useOfficeSheetStore();
    await store.load('sheet-1');

    store.toggleBold('Sheet1', ['A1']);
    store.mergeRange('Sheet1', 'A1:C1');
    store.fillCells('Sheet1', 'A1', ['A2']);
    await vi.advanceTimersByTimeAsync(500);

    expect(mockSaveCells).not.toHaveBeenCalled();
    store.stopPolling();
  });
});
