import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';
import { createPinia, setActivePinia } from 'pinia';

const mockGetSheet = vi.fn();
const mockSaveCells = vi.fn();
const mockRecalc = vi.fn();
const mockExportWorkbook = vi.fn();
const mockImportWorkbook = vi.fn();
const mockAcquireLease = vi.fn();
const mockReleaseLease = vi.fn();
const mockPresence = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeSheetApi: {
    getSheet: (...args: unknown[]) => mockGetSheet(...args),
    createSheet: vi.fn(),
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
  OfficeDocConflictError: class extends Error {},
}));

vi.mock('vue-router', () => ({
  useRoute: () => ({ params: { id: 'sheet-1' } }),
}));

import OfficeSheetEditor from '../src/views/OfficeSheetEditor.vue';

const tFallback = (key: string) => key;

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
      sheets: [{ name: 'Sheet1', cells: { A1: { v: 42 }, B1: { f: '=A1+1', v: 43 } } }],
      active_sheet: 'Sheet1',
    },
    version_no: 1,
    access: 'owner',
    lease: { held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null },
    ...overrides,
  };
}

function mountEditor() {
  return mount(OfficeSheetEditor, {
    global: {
      mocks: { $t: tFallback },
      stubs: { 'router-link': true },
    },
  });
}

describe('OfficeSheetEditor', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
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

  it('loads the sheet and renders the grid surface with the stored values', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    const wrapper = mountEditor();
    await flushPromises();

    expect(mockGetSheet).toHaveBeenCalledWith('sheet-1');
    expect(wrapper.find('[data-testid="office-sheet-editor"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-sheet-grid"]').exists()).toBe(true);
    expect(wrapper.text()).toContain('42');
    expect(wrapper.text()).toContain('43');
    wrapper.unmount();
  });

  it('renders one tab per sheet', async () => {
    mockGetSheet.mockResolvedValue(
      sheetView({
        workbook: {
          sheets: [
            { name: 'Sheet1', cells: {} },
            { name: 'Sheet2', cells: {} },
          ],
          active_sheet: 'Sheet1',
        },
      }),
    );
    const wrapper = mountEditor();
    await flushPromises();

    const tabs = wrapper.findAll('[data-testid="office-sheet-tab"]');
    expect(tabs).toHaveLength(2);
    wrapper.unmount();
  });

  it('shows the view-only banner and disables the import button for a view-only access', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ access: 'view' }));
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-sheet-view-only-banner"]').exists()).toBe(true);
    expect(
      (wrapper.find('[data-testid="office-sheet-import-btn"]').element as HTMLButtonElement).disabled,
    ).toBe(true);
    wrapper.unmount();
  });

  it('shows the lease banner with a take-over action when another editor holds it', async () => {
    const someoneElseLease = {
      held: true,
      holder_user_id: 'someone-else',
      is_self: false,
      granted: false,
      expires_at: '2026-07-28T00:05:00Z',
    };
    mockGetSheet.mockResolvedValue(sheetView({ lease: someoneElseLease }));
    mockAcquireLease.mockResolvedValue(someoneElseLease);
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-sheet-lease-banner"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-sheet-lease-takeover"]').exists()).toBe(true);
    wrapper.unmount();
  });

  it('clicking a cell selects it and populates the formula bar with its raw content', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    const wrapper = mountEditor();
    await flushPromises();

    const cells = wrapper.findAll('[data-testid="office-sheet-cell"]');
    // Row 1, column 2 (B1) — the formula cell.
    await cells[1].trigger('click');

    const formulaInput = wrapper.find('[data-testid="office-sheet-formula-bar-input"]')
      .element as HTMLInputElement;
    expect(formulaInput.value).toBe('=A1+1');
    wrapper.unmount();
  });

  it('the recalc button calls recalcAll and applies the returned deltas', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    mockRecalc.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!A1': 100 } });
    const wrapper = mountEditor();
    await flushPromises();

    await wrapper.find('[data-testid="office-sheet-recalc-btn"]').trigger('click');
    await flushPromises();

    expect(mockRecalc).toHaveBeenCalledWith('sheet-1');
    expect(wrapper.text()).toContain('100');
    wrapper.unmount();
  });
});
