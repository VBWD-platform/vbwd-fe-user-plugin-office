import { describe, it, expect, beforeEach, vi } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';

const mockDownloadDocument = vi.fn();

vi.mock('../src/api/officeApi', () => ({
  officeApi: {
    downloadDocument: (...args: unknown[]) => mockDownloadDocument(...args),
    contentUrl: (id: string) => `/api/v1/office/documents/${id}/content`,
  },
}));

import OfficePreviewPane from '../src/components/OfficePreviewPane.vue';

/**
 * The bug this pins: images and PDFs were rendered with a bare
 * `:src="contentUrl(id)"`. A browser does NOT attach the Authorization header to
 * an <img>/<iframe> request, so the content route answered 401 and the pane
 * showed an empty box — reported as "click a file and something flashes and
 * disappears".
 *
 * The assertion is therefore about the TRANSPORT, not the markup: the component
 * must FETCH the bytes (which carries auth) rather than delegate loading to the
 * browser. Checking only that an <img> exists would have passed throughout the
 * entire time the feature was broken.
 */
const IMAGE_NODE = {
  id: 'n-img', name: 'photo.jpeg', kind: 'document', mime_type: 'image/jpeg',
  parent_id: null, trashed_at: null, created_at: null, updated_at: null,
} as never;
const PDF_NODE = { ...(IMAGE_NODE as object), id: 'n-pdf', name: 'a.pdf', mime_type: 'application/pdf' } as never;
const TEXT_NODE = { ...(IMAGE_NODE as object), id: 'n-txt', name: 'a.txt', mime_type: 'text/plain' } as never;

function mountPane(node: never) {
  return mount(OfficePreviewPane, {
    props: { node },
    global: { mocks: { $t: (key: string) => key } },
  });
}

describe('OfficePreviewPane — media is fetched with auth, never left to the browser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('URL', {
      createObjectURL: vi.fn(() => 'blob:fake-object-url'),
      revokeObjectURL: vi.fn(),
    });
  });

  it('fetches an image through the API client instead of a bare src', async () => {
    mockDownloadDocument.mockResolvedValue({ blob: new Blob(['x'], { type: 'image/jpeg' }) });
    const wrapper = mountPane(IMAGE_NODE);
    await flushPromises();

    expect(mockDownloadDocument).toHaveBeenCalledWith('n-img', 'photo.jpeg');
    const src = wrapper.get('[data-testid="office-preview-image"]').attributes('src');
    // An object URL, NOT the authenticated API path a browser cannot authorise.
    expect(src).toBe('blob:fake-object-url');
    expect(src).not.toContain('/api/v1/office/documents');
  });

  it('fetches a PDF the same way', async () => {
    mockDownloadDocument.mockResolvedValue({ blob: new Blob(['%PDF-'], { type: 'application/pdf' }) });
    const wrapper = mountPane(PDF_NODE);
    await flushPromises();

    expect(mockDownloadDocument).toHaveBeenCalledWith('n-pdf', 'a.pdf');
    expect(wrapper.get('[data-testid="office-preview-pdf"]').attributes('src')).toBe(
      'blob:fake-object-url',
    );
  });

  it('still renders text content as text', async () => {
    mockDownloadDocument.mockResolvedValue({ blob: new Blob(['hello there'], { type: 'text/plain' }) });
    const wrapper = mountPane(TEXT_NODE);
    await flushPromises();

    expect(wrapper.get('[data-testid="office-preview-text"]').text()).toContain('hello there');
  });

  it('says so when the fetch fails rather than showing an empty pane', async () => {
    mockDownloadDocument.mockRejectedValue(new Error('401'));
    const wrapper = mountPane(IMAGE_NODE);
    await flushPromises();

    expect(wrapper.find('[data-testid="office-preview-error"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="office-preview-image"]').exists()).toBe(false);
  });

  it('revokes the object URL on unmount so blobs are not pinned in memory', async () => {
    mockDownloadDocument.mockResolvedValue({ blob: new Blob(['x'], { type: 'image/jpeg' }) });
    const wrapper = mountPane(IMAGE_NODE);
    await flushPromises();

    wrapper.unmount();
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:fake-object-url');
  });
});
