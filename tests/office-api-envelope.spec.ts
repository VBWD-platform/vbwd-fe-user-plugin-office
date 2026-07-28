import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mocked at the TRANSPORT level (global.fetch), not at the officeApi method
// level — a method-level mock would bypass the exact bug this file guards
// against: the backend wraps every list response in an envelope
// (`jsonify({"items": [...]})`, `routes.py`), and a client method that
// forgets to unwrap it hands the STORE `{items: [...]}` instead of an array,
// which then throws on the first `.filter(...)` / `.map(...)` call — this
// was invisible to method-level-mocked tests because they returned a bare
// array directly, never exercising the real HTTP envelope at all.
import { officeApi } from '../src/api/officeApi';

function jsonResponse(body: unknown, status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  } as Response;
}

describe('officeApi — response envelope unwrapping', () => {
  beforeEach(() => {
    global.fetch = vi.fn();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('listNodes() returns a plain array, not the {items: [...]} envelope', async () => {
    const node = {
      id: 'n1',
      parent_id: null,
      kind: 'document',
      name: 'a.txt',
      trashed_at: null,
      created_at: '2026-07-28T00:00:00Z',
      updated_at: '2026-07-28T00:00:00Z',
    };
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ items: [node] }),
    );

    const result = await officeApi.listNodes(null);

    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(1);
    expect(result[0].id).toBe('n1');
    // The historical bug: `store.nodes.filter(...)` would throw here because
    // `store.nodes` was the envelope object, not an array.
    expect(() => result.filter((n) => n.kind === 'document')).not.toThrow();
  });

  it('listVersions() returns a plain array, not the {items: [...]} envelope', async () => {
    const version = {
      id: 'v1',
      version_no: 1,
      size_bytes: 5,
      sha256: 'abc',
      created_by_user_id: 'u1',
      created_at: '2026-07-28T00:00:00Z',
    };
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ items: [version] }),
    );

    const result = await officeApi.listVersions('doc-1');

    expect(Array.isArray(result)).toBe(true);
    expect(result[0].version_no).toBe(1);
  });

  it('restoreVersion() sends {version_no}, matching the backend contract exactly', async () => {
    const fetchMock = global.fetch as ReturnType<typeof vi.fn>;
    fetchMock.mockResolvedValue(
      jsonResponse({ id: 'v2', version_no: 2, size_bytes: 5, sha256: 'x', created_by_user_id: 'u1', created_at: '2026-07-28T00:00:00Z' }),
    );

    await officeApi.restoreVersion('doc-1', 1);

    const [, requestInit] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(JSON.parse(requestInit.body as string)).toEqual({ version_no: 1 });
  });

  it('getUsage() and getDocument() stay bare objects (no envelope on single-resource routes)', async () => {
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ bytes_used: 10, bytes_quota: 100 }),
    );

    const usage = await officeApi.getUsage();

    expect(usage).toEqual({ bytes_used: 10, bytes_quota: 100 });
  });

  it('listNodes() flattens document.{mime_type,size_bytes,doc_type} onto the node', async () => {
    // NodeSummary.to_dict() (routes.py) nests document facts under a
    // `document` sub-object; SpaceHome.vue and OfficePreviewPane.vue read
    // `node.mime_type` / `node.size_bytes` FLAT. Left un-normalised, every
    // document in the list silently shows "0 B" and never previews inline —
    // invisible to a method-mocked test for the same reason as the envelope
    // bug: the mock bypassed the real wire shape entirely.
    const documentNode = {
      id: 'n1',
      parent_id: null,
      kind: 'document',
      name: 'report.pdf',
      trashed_at: null,
      created_at: '2026-07-28T00:00:00Z',
      updated_at: '2026-07-28T00:00:00Z',
      document: { id: 'doc1', doc_type: 'file', mime_type: 'application/pdf', size_bytes: 255 },
    };
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ items: [documentNode] }),
    );

    const [result] = await officeApi.listNodes(null);

    expect(result.size_bytes).toBe(255);
    expect(result.mime_type).toBe('application/pdf');
    expect(result.doc_type).toBe('file');
  });
});
