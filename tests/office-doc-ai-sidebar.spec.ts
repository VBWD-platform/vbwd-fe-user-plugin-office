import { describe, it, expect } from 'vitest';
import { mount } from '@vue/test-utils';
import OfficeDocAiSidebar from '../src/components/OfficeDocAiSidebar.vue';

const tFallback = (key: string) => key;

function mountSidebar(props: Record<string, unknown> = {}) {
  return mount(OfficeDocAiSidebar, {
    props: {
      aiEnabled: true,
      canToggle: true,
      hasSelection: false,
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

describe('OfficeDocAiSidebar', () => {
  it('shows the disabled hint and hides capabilities/freeform prompt when the AI helper is off', () => {
    const wrapper = mountSidebar({ aiEnabled: false });
    expect(wrapper.find('[data-testid="office-doc-ai-disabled-hint"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-doc-ai-capability-summarize"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="office-doc-ai-prompt-input"]').exists()).toBe(false);
  });

  it('renders every preset capability button alongside the freeform prompt', () => {
    const wrapper = mountSidebar();
    for (const id of [
      'continue_writing',
      'rewrite_shorter',
      'rewrite_longer',
      'rewrite_formal',
      'rewrite_plain',
      'summarize',
      'fix_grammar',
      'translate',
      'outline',
    ]) {
      expect(wrapper.find(`[data-testid="office-doc-ai-capability-${id}"]`).exists()).toBe(true);
    }
    expect(wrapper.find('[data-testid="office-doc-ai-prompt-input"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-doc-ai-prompt-run"]').exists()).toBe(true);
  });

  it('disables the freeform Run button until a prompt is typed, and emits it when clicked', async () => {
    const wrapper = mountSidebar();
    expect(wrapper.get('[data-testid="office-doc-ai-prompt-run"]').attributes('disabled')).toBeDefined();

    await wrapper.get('[data-testid="office-doc-ai-prompt-input"]').setValue('rewrite this as three bullet points');
    expect(wrapper.get('[data-testid="office-doc-ai-prompt-run"]').attributes('disabled')).toBeUndefined();

    await wrapper.get('[data-testid="office-doc-ai-prompt-run"]').trigger('click');
    expect(wrapper.emitted('run-freeform')?.[0]).toEqual(['rewrite this as three bullet points']);
  });

  it('a whitespace-only prompt never emits run-freeform', async () => {
    const wrapper = mountSidebar();
    await wrapper.get('[data-testid="office-doc-ai-prompt-input"]').setValue('   ');
    expect(wrapper.get('[data-testid="office-doc-ai-prompt-run"]').attributes('disabled')).toBeDefined();
  });

  it('the freeform Run button is disabled while a request is running', () => {
    const wrapper = mountSidebar({ running: true });
    expect(wrapper.find('[data-testid="office-doc-ai-running"]').exists()).toBe(true);
  });

  it('a preset capability click still emits run-capability, unaffected by the freeform addition', async () => {
    const wrapper = mountSidebar();
    await wrapper.get('[data-testid="office-doc-ai-capability-continue_writing"]').trigger('click');
    expect(wrapper.emitted('run-capability')?.[0]).toEqual(['continue_writing']);
  });

  it('an error shows the translated error message', () => {
    const wrapper = mountSidebar({ error: 'office.ai.forbidden' });
    expect(wrapper.get('[data-testid="office-doc-ai-error"]').text()).toBe('office.ai.forbidden');
  });

  it('shows the proposed text and emits accept/discard', async () => {
    const wrapper = mountSidebar({
      proposal: { capability: 'freeform', proposedText: '- point one\n- point two' },
    });
    expect(wrapper.get('[data-testid="office-doc-ai-proposal"]').text()).toContain('point one');
    await wrapper.get('[data-testid="office-doc-ai-accept"]').trigger('click');
    expect(wrapper.emitted('accept')).toBeTruthy();
    await wrapper.get('[data-testid="office-doc-ai-discard"]').trigger('click');
    expect(wrapper.emitted('discard')).toBeTruthy();
  });

  it('emits toggle-enabled from the checkbox', async () => {
    const wrapper = mountSidebar();
    const checkbox = wrapper.get('[data-testid="office-doc-ai-enabled-toggle"]');
    await checkbox.setValue(false);
    expect(wrapper.emitted('toggle-enabled')?.[0]).toEqual([false]);
  });
});

// Guard against `useI18n()` sneaking into script scope without a mock — this
// component relies entirely on the template's `$t`, so no `vue-i18n` mock is
// registered here (see office-doc-editor.spec.ts's note on this convention).
describe('OfficeDocAiSidebar — no unexpected vue-i18n dependency', () => {
  it('mounts without a vue-i18n plugin/mock installed', () => {
    expect(() => mountSidebar()).not.toThrow();
  });
});
