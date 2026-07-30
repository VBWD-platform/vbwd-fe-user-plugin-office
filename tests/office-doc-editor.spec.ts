import { describe, it, expect, vi, beforeEach } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';
import { createPinia, setActivePinia } from 'pinia';

const mockGetDoc = vi.fn();
const mockAcquireLease = vi.fn();
const mockReleaseLease = vi.fn();
const mockPresence = vi.fn();
const mockListVersions = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    listVersions: (...args: unknown[]) => mockListVersions(...args),
    restoreVersion: vi.fn(),
  },
  officeDocApi: {
    getDoc: (...args: unknown[]) => mockGetDoc(...args),
    createDoc: vi.fn(),
    saveDoc: vi.fn(),
    acquireLease: (...args: unknown[]) => mockAcquireLease(...args),
    releaseLease: (...args: unknown[]) => mockReleaseLease(...args),
    presence: (...args: unknown[]) => mockPresence(...args),
    runAiCapability: vi.fn(),
  },
  OfficeDocConflictError: class extends Error {},
  OfficeAiForbiddenError: class extends Error {},
  OfficeAiBudgetExceededError: class extends Error {},
}));

// OfficeDocEditor calls `useI18n()` for the editor placeholder, which needs the
// string in script scope where the template-only `$t` is unavailable. The real
// plugin is installed by the app; mount tests supply it directly.
vi.mock('vue-i18n', () => ({
  useI18n: () => ({ t: (key: string) => key }),
}));

vi.mock('vue-router', () => ({
  useRoute: () => ({ params: { id: 'doc-1' } }),
}));

import OfficeDocEditor from '../src/views/OfficeDocEditor.vue';

const tFallback = (key: string) => key;

const EMPTY_CONTENT = { type: 'doc', content: [{ type: 'paragraph' }] };

function baseDocView(overrides: Record<string, unknown> = {}) {
  return {
    id: 'doc-1',
    parent_id: null,
    kind: 'document',
    name: 'My Doc',
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

function mountEditor() {
  return mount(OfficeDocEditor, {
    global: {
      mocks: { $t: tFallback },
      stubs: { 'router-link': true },
    },
  });
}

describe('OfficeDocEditor', () => {
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

  it('loads the document and renders the editor surface', async () => {
    mockGetDoc.mockResolvedValue(baseDocView());
    const wrapper = mountEditor();
    await flushPromises();

    expect(mockGetDoc).toHaveBeenCalledWith('doc-1');
    expect(wrapper.find('[data-testid="office-doc-editor"]').exists()).toBe(true);
    wrapper.unmount();
  });

  it('shows the view-only banner and hides the toolbar for a view-only access', async () => {
    mockGetDoc.mockResolvedValue(baseDocView({ access: 'view' }));
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-doc-view-only-banner"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-doc-lease-banner"]').exists()).toBe(false);
    wrapper.unmount();
  });

  it('shows the lease banner with a take-over action when another editor holds it', async () => {
    const someoneElseLease = {
      held: true,
      holder_user_id: 'someone-else',
      is_self: false,
      granted: false,
      expires_at: '2026-07-27T00:05:00Z',
    };
    mockGetDoc.mockResolvedValue(baseDocView({ lease: someoneElseLease }));
    // The store's own acquire-on-load call must ALSO report "not granted" —
    // otherwise the store would overwrite the loaded lease with a
    // self-granted one, defeating the scenario this test sets up.
    mockAcquireLease.mockResolvedValue(someoneElseLease);
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-doc-lease-banner"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-doc-lease-takeover"]').exists()).toBe(true);
    wrapper.unmount();
  });

  it('toggling the AI panel reveals the AI sidebar testid', async () => {
    mockGetDoc.mockResolvedValue(baseDocView());
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.find('[data-testid="office-doc-ai-sidebar"]').exists()).toBe(false);
    await wrapper.find('[data-testid="office-doc-ai-toggle-panel"]').trigger('click');
    expect(wrapper.find('[data-testid="office-doc-ai-sidebar"]').exists()).toBe(true);
    wrapper.unmount();
  });
});

describe('OfficeDocEditor — i18n', () => {
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

  it('translates the save-status label instead of rendering the raw i18n key', async () => {
    // The suite's default $t is the identity function, which cannot tell a
    // translated string from an untranslated key — which is exactly how
    // `{{ saveStatusLabel }}` shipped showing a literal "office.doc.saved" to
    // the user. This mount uses a MARKING translator, so the assertion proves
    // the value actually went through $t rather than being interpolated raw.
    mockGetDoc.mockResolvedValue(baseDocView());

    const wrapper = mount(OfficeDocEditor, {
      global: {
        mocks: { $t: (key: string) => `translated:${key}` },
        stubs: { 'router-link': true },
      },
    });
    await flushPromises();

    expect(wrapper.get('[data-testid="office-doc-save-status"]').text()).toBe(
      'translated:office.doc.saved',
    );
  });
});

/**
 * The reported symptom was "I don't see a cursor — I cannot edit". The document
 * was fine: it was an EMPTY doc with no caret and no placeholder, which is
 * visually indistinguishable from a read-only one. These pin both halves of the
 * fix, plus the unhelpful raw-404 message that sent the diagnosis the wrong way.
 */
describe('OfficeDocEditor — an empty document must look editable', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.clearAllMocks();
    mockAcquireLease.mockResolvedValue({
      held: true, holder_user_id: 'me', is_self: true, granted: true,
      expires_at: '2026-07-30T00:01:30Z',
    });
    mockPresence.mockResolvedValue({
      held: false, holder_user_id: null, is_self: false, granted: true, expires_at: null,
    });
  });

  it('autofocuses the editor when the document is editable', async () => {
    mockGetDoc.mockResolvedValue(baseDocView());
    const wrapper = mountEditor();
    await flushPromises();

    // Tiptap resolves `autofocus` into its own options; asserting on the option
    // keeps this independent of jsdom's incomplete focus emulation.
    expect((wrapper.vm as unknown as { editor: { options: { autofocus: unknown } } }).editor
      .options.autofocus).toBe('end');
  });

  it('does NOT autofocus a document the user cannot edit', async () => {
    mockGetDoc.mockResolvedValue(baseDocView({ access: 'view' }));
    const wrapper = mountEditor();
    await flushPromises();

    expect((wrapper.vm as unknown as { editor: { options: { autofocus: unknown } } }).editor
      .options.autofocus).toBe(false);
  });

  it('registers a placeholder so an empty document says it is writable', async () => {
    mockGetDoc.mockResolvedValue(baseDocView());
    const wrapper = mountEditor();
    await flushPromises();

    const names = (wrapper.vm as unknown as {
      editor: { extensionManager: { extensions: Array<{ name: string }> } };
    }).editor.extensionManager.extensions.map((extension) => extension.name);
    expect(names).toContain('placeholder');
  });

  it('renders a human message for a 404 rather than the raw request error', async () => {
    mockGetDoc.mockRejectedValue(
      new Error('GET /api/v1/office/docs/ccc16883 failed: 404'),
    );
    const wrapper = mountEditor();
    await flushPromises();

    const text = wrapper.get('[data-testid="office-doc-load-error"]').text();
    // The identity $t mock returns the key, which is what proves the store
    // mapped the raw error onto a translatable message at all.
    expect(text).toBe('office.doc.notFound');
    expect(text).not.toContain('404');
  });

  it('passes a non-404 failure through unchanged rather than mislabelling it', async () => {
    mockGetDoc.mockRejectedValue(new Error('NetworkError: connection refused'));
    const wrapper = mountEditor();
    await flushPromises();

    expect(wrapper.get('[data-testid="office-doc-load-error"]').text()).toContain(
      'connection refused',
    );
  });
});
