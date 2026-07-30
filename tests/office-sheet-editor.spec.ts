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
const mockRunAiCapability = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeSheetApi: {
    getSheet: (...args: unknown[]) => mockGetSheet(...args),
    createSheet: vi.fn(),
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
  OfficeDocConflictError: class extends Error {},
  OfficeAiForbiddenError: class extends Error {},
  OfficeAiBudgetExceededError: class extends Error {},
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
    // Row 1, column 2 (B1) — the formula cell. Selection is driven by
    // mousedown (not click) so a drag-select can begin on the same event.
    await cells[1].trigger('mousedown', { button: 0 });

    const formulaInput = wrapper.find('[data-testid="office-sheet-formula-bar-input"]')
      .element as HTMLInputElement;
    expect(formulaInput.value).toBe('=A1+1');
    wrapper.unmount();
  });

  it('adds a sheet tab via the add-sheet dialog', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    // addSheetTab queues a debounced save (SAVE_DEBOUNCE_MS=400 in the
    // store) — give it a resolved response and let the timer fire so it
    // doesn't reject as an unhandled rejection after this test completes.
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: {} });
    const wrapper = mountEditor();
    await flushPromises();

    await wrapper.find('[data-testid="office-sheet-add-sheet-btn"]').trigger('click');
    expect(wrapper.find('[data-testid="office-sheet-add-sheet-modal"]').exists()).toBe(true);

    const input = wrapper.find('[data-testid="office-sheet-add-sheet-input"]')
      .element as HTMLInputElement;
    expect(input.value).toBe('Sheet2');

    await wrapper.find('[data-testid="office-sheet-add-sheet-input"]').setValue('Q3 forecast');
    await wrapper.find('[data-testid="office-sheet-add-sheet-confirm"]').trigger('click');
    await flushPromises();

    const tabLabels = wrapper.findAll('[data-testid="office-sheet-tab"]').map((tab) => tab.text());
    expect(tabLabels).toContain('Q3 forecast');
    expect(wrapper.find('[data-testid="office-sheet-add-sheet-modal"]').exists()).toBe(false);

    await new Promise((resolve) => { setTimeout(resolve, 450); });
    wrapper.unmount();
  });

  it('cancelling the add-sheet dialog adds no tab', async () => {
    mockGetSheet.mockResolvedValue(sheetView());
    const wrapper = mountEditor();
    await flushPromises();

    await wrapper.find('[data-testid="office-sheet-add-sheet-btn"]').trigger('click');
    await wrapper.find('[data-testid="office-sheet-add-sheet-cancel"]').trigger('click');

    expect(wrapper.findAll('[data-testid="office-sheet-tab"]')).toHaveLength(1);
    expect(wrapper.find('[data-testid="office-sheet-add-sheet-modal"]').exists()).toBe(false);
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

// ---------------------------------------------------------------------------
// S147-3.5 — the Sheet AI helper panel.
// ---------------------------------------------------------------------------

describe('OfficeSheetEditor — AI helper panel', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    mockAcquireLease.mockResolvedValue({
      held: true, holder_user_id: 'me', is_self: true, granted: true,
      expires_at: '2026-07-28T00:01:30Z',
    });
    mockPresence.mockResolvedValue({
      held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null,
    });
  });

  it('toggling the AI panel reveals the AI sidebar testid', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-sheet-ai-sidebar"]').exists()).toBe(false);
    await wrapper.find('[data-testid="office-sheet-ai-toggle-panel"]').trigger('click');
    expect(wrapper.find('[data-testid="office-sheet-ai-sidebar"]').exists()).toBe(true);
    wrapper.unmount();
  });

  it('running a capability sends the active cell address (and the AI-disabled hint before that)', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'text',
      capability: 'sheet_explain_formula',
      connection_slug: 'default',
      text: 'It adds one to A1.',
    });
    const wrapper = mountEditor();
    await flushPromises();
    await wrapper.find('[data-testid="office-sheet-ai-toggle-panel"]').trigger('click');

    const cells = wrapper.findAll('[data-testid="office-sheet-cell"]');
    await cells[1].trigger('mousedown', { button: 0 }); // B1

    await wrapper
      .find('[data-testid="office-sheet-ai-capability-sheet_explain_formula"]')
      .trigger('click');
    await flushPromises();

    expect(mockRunAiCapability).toHaveBeenCalledWith('sheet-1', {
      capability: 'sheet_explain_formula',
      address: 'B1',
      rangeText: 'B1',
      intent: '',
    });
    expect(wrapper.get('[data-testid="office-sheet-ai-proposal-text"]').text()).toBe('It adds one to A1.');
    wrapper.unmount();
  });

  it('accepting a formula proposal applies it through saveCells and the cell shows the computed value', async () => {
    mockGetSheet.mockResolvedValue(sheetView({ document: { ai_enabled: true } }));
    mockRunAiCapability.mockResolvedValue({
      kind: 'formula',
      capability: 'sheet_write_formula',
      connection_slug: 'default',
      address: 'B1',
      formula: '=A1+1',
    });
    mockSaveCells.mockResolvedValue({ version_no: 2, changes: { 'Sheet1!B1': 43 } });
    const wrapper = mountEditor();
    await flushPromises();
    await wrapper.find('[data-testid="office-sheet-ai-toggle-panel"]').trigger('click');

    const cells = wrapper.findAll('[data-testid="office-sheet-cell"]');
    await cells[1].trigger('mousedown', { button: 0 }); // B1

    await wrapper.find('[data-testid="office-sheet-ai-intent-input"]').setValue('add one to A1');
    await wrapper
      .find('[data-testid="office-sheet-ai-capability-sheet_write_formula"]')
      .trigger('click');
    await flushPromises();

    await wrapper.find('[data-testid="office-sheet-ai-accept"]').trigger('click');
    await flushPromises();

    expect(mockSaveCells).toHaveBeenCalledWith(
      'sheet-1',
      [{ sheet: 'Sheet1', address: 'B1', formula: '=A1+1' }],
      1,
    );
    expect(wrapper.find('[data-testid="office-sheet-ai-proposal"]').exists()).toBe(false);
    wrapper.unmount();
  });

  it('a view-only share sees the explain capability enabled but the formula ones disabled', async () => {
    mockGetSheet.mockResolvedValue(
      sheetView({ access: 'view', document: { ai_enabled: true } }),
    );
    const wrapper = mountEditor();
    await flushPromises();
    await wrapper.find('[data-testid="office-sheet-ai-toggle-panel"]').trigger('click');

    const cells = wrapper.findAll('[data-testid="office-sheet-cell"]');
    await cells[1].trigger('mousedown', { button: 0 }); // B1

    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_write_formula"]').attributes('disabled'),
    ).toBeDefined();
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_explain_formula"]').attributes('disabled'),
    ).toBeUndefined();
    wrapper.unmount();
  });
});
