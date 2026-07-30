import { describe, it, expect } from 'vitest';
import { mount } from '@vue/test-utils';
import OfficeSheetAiSidebar from '../src/components/OfficeSheetAiSidebar.vue';

const tFallback = (key: string) => key;

function mountSidebar(props: Record<string, unknown> = {}) {
  return mount(OfficeSheetAiSidebar, {
    props: {
      aiEnabled: true,
      canToggle: true,
      canProposeFormula: true,
      hasActiveCell: true,
      running: false,
      error: null,
      proposal: null,
      ...props,
    },
    global: {
      mocks: { $t: tFallback },
    },
  });
}

describe('OfficeSheetAiSidebar', () => {
  it('shows the disabled hint and hides capabilities when the AI helper is off', () => {
    const wrapper = mountSidebar({ aiEnabled: false });
    expect(wrapper.find('[data-testid="office-sheet-ai-disabled-hint"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-sheet-ai-capability-sheet_explain_formula"]').exists()).toBe(false);
  });

  it('shows a select-a-cell hint before any cell is active', () => {
    const wrapper = mountSidebar({ hasActiveCell: false });
    expect(wrapper.find('[data-testid="office-sheet-ai-select-cell-hint"]').exists()).toBe(true);
  });

  it('renders all four capability buttons once a cell is active', () => {
    const wrapper = mountSidebar();
    for (const id of [
      'sheet_write_formula',
      'sheet_fix_error',
      'sheet_explain_formula',
      'sheet_summarize_range',
    ]) {
      expect(wrapper.find(`[data-testid="office-sheet-ai-capability-${id}"]`).exists()).toBe(true);
    }
  });

  it('disables the formula-producing capabilities without edit access', () => {
    const wrapper = mountSidebar({ canProposeFormula: false });
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_write_formula"]').attributes('disabled'),
    ).toBeDefined();
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_fix_error"]').attributes('disabled'),
    ).toBeDefined();
    // Read-only capabilities stay enabled — a view share may explain/summarise.
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_explain_formula"]').attributes('disabled'),
    ).toBeUndefined();
  });

  it('disables "write formula" until an intent is typed, and emits it when clicked', async () => {
    const wrapper = mountSidebar();
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_write_formula"]').attributes('disabled'),
    ).toBeDefined();

    await wrapper.get('[data-testid="office-sheet-ai-intent-input"]').setValue('total column A');
    expect(
      wrapper.get('[data-testid="office-sheet-ai-capability-sheet_write_formula"]').attributes('disabled'),
    ).toBeUndefined();

    await wrapper.get('[data-testid="office-sheet-ai-capability-sheet_write_formula"]').trigger('click');
    expect(wrapper.emitted('run-capability')?.[0]).toEqual(['sheet_write_formula', 'total column A']);
  });

  it('running shows the running indicator', () => {
    const wrapper = mountSidebar({ running: true });
    expect(wrapper.find('[data-testid="office-sheet-ai-running"]').exists()).toBe(true);
  });

  it('an error shows the translated error message', () => {
    const wrapper = mountSidebar({ error: 'office.sheet.ai.forbidden' });
    expect(wrapper.get('[data-testid="office-sheet-ai-error"]').text()).toBe('office.sheet.ai.forbidden');
  });

  it('shows the formula as TEXT before it is accepted', () => {
    const wrapper = mountSidebar({
      proposal: {
        kind: 'formula',
        capability: 'sheet_write_formula',
        connection_slug: 'default',
        address: 'B1',
        formula: '=SUM(A1:A2)',
      },
    });
    const formulaEl = wrapper.get('[data-testid="office-sheet-ai-proposal-formula"]');
    expect(formulaEl.text()).toContain('B1');
    expect(formulaEl.text()).toContain('=SUM(A1:A2)');
  });

  it('shows a text proposal as prose', () => {
    const wrapper = mountSidebar({
      proposal: {
        kind: 'text',
        capability: 'sheet_summarize_range',
        connection_slug: 'default',
        text: 'A1:A2 sum to 30.',
      },
    });
    expect(wrapper.get('[data-testid="office-sheet-ai-proposal-text"]').text()).toBe('A1:A2 sum to 30.');
  });

  it('emits accept/discard', async () => {
    const wrapper = mountSidebar({
      proposal: {
        kind: 'text',
        capability: 'sheet_summarize_range',
        connection_slug: 'default',
        text: 'A1:A2 sum to 30.',
      },
    });
    await wrapper.get('[data-testid="office-sheet-ai-accept"]').trigger('click');
    expect(wrapper.emitted('accept')).toBeTruthy();
    await wrapper.get('[data-testid="office-sheet-ai-discard"]').trigger('click');
    expect(wrapper.emitted('discard')).toBeTruthy();
  });

  it('emits toggle-enabled from the checkbox', async () => {
    const wrapper = mountSidebar();
    const checkbox = wrapper.get('[data-testid="office-sheet-ai-enabled-toggle"]');
    await checkbox.setValue(false);
    expect(wrapper.emitted('toggle-enabled')?.[0]).toEqual([false]);
  });

  it('the toggle is disabled for a non-owner', () => {
    const wrapper = mountSidebar({ canToggle: false });
    expect(wrapper.get('[data-testid="office-sheet-ai-enabled-toggle"]').attributes('disabled')).toBeDefined();
  });
});

// Guard against `useI18n()` sneaking into script scope without a mock — this
// component currently relies entirely on the template's `$t`, so no
// `vue-i18n` mock is registered here (see office-doc-editor.spec.ts's note).
describe('OfficeSheetAiSidebar — no unexpected vue-i18n dependency', () => {
  it('mounts without a vue-i18n plugin/mock installed', () => {
    expect(() => mountSidebar()).not.toThrow();
  });
});
