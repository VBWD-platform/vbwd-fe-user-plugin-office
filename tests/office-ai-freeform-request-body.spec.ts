import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mocked at the TRANSPORT level (global.fetch), not at the officeApi method
// level — see office-api-envelope.spec.ts's docstring for why. This file
// guards the FE mirror of the backend's raw-prompt-rejection control (S147
// free-text helper): a `prompt` field must be sent ONLY for the two
// freeform capabilities, never smuggled onto a preset capability's request
// body, where the backend would 400 it anyway — the client should not even
// try.
import { officeDocApi, officeSheetApi } from '../src/api/officeApi';

function jsonResponse(body: unknown, status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  } as Response;
}

function lastRequestBody(fetchMock: ReturnType<typeof vi.fn>): Record<string, unknown> {
  const [, init] = fetchMock.mock.calls[fetchMock.mock.calls.length - 1];
  return JSON.parse((init as RequestInit).body as string);
}

describe('officeDocApi.runAiCapability — prompt field scoping', () => {
  beforeEach(() => {
    global.fetch = vi.fn();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('omits "prompt" entirely for a preset capability', async () => {
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ capability: 'summarize', connection_slug: 'default', proposed_text: 'ok' }),
    );

    await officeDocApi.runAiCapability('doc-1', {
      capability: 'summarize',
      selectionText: 'hello world',
    });

    const body = lastRequestBody(global.fetch as ReturnType<typeof vi.fn>);
    expect('prompt' in body).toBe(false);
  });

  it('sends "prompt" for the freeform capability', async () => {
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ capability: 'freeform', connection_slug: 'default', proposed_text: 'ok' }),
    );

    await officeDocApi.runAiCapability('doc-1', {
      capability: 'freeform',
      selectionText: 'hello world',
      prompt: 'rewrite this as three bullet points',
    });

    const body = lastRequestBody(global.fetch as ReturnType<typeof vi.fn>);
    expect(body.prompt).toBe('rewrite this as three bullet points');
  });
});

describe('officeSheetApi.runAiCapability — prompt field scoping', () => {
  beforeEach(() => {
    global.fetch = vi.fn();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('omits "prompt" entirely for a preset capability', async () => {
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ kind: 'text', capability: 'sheet_explain_formula', connection_slug: 'default', text: 'ok' }),
    );

    await officeSheetApi.runAiCapability('sheet-1', {
      capability: 'sheet_explain_formula',
      address: 'A1',
    });

    const body = lastRequestBody(global.fetch as ReturnType<typeof vi.fn>);
    expect('prompt' in body).toBe(false);
  });

  it('sends "prompt" for the sheet_freeform capability', async () => {
    (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValue(
      jsonResponse({ kind: 'formula', capability: 'sheet_freeform', connection_slug: 'default', formula: '=B1*2' }),
    );

    await officeSheetApi.runAiCapability('sheet-1', {
      capability: 'sheet_freeform',
      address: 'A1',
      prompt: 'add a column that is column B times 2',
    });

    const body = lastRequestBody(global.fetch as ReturnType<typeof vi.fn>);
    expect(body.prompt).toBe('add a column that is column B times 2');
  });
});
