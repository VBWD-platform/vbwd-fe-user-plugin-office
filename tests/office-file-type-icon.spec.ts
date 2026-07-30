import { describe, it, expect } from 'vitest';
import { hasIcon } from 'vbwd-view-component';
import { fileTypeIcon } from '../src/utils/fileTypeIcon';

describe('fileTypeIcon', () => {
  it('maps a folder to the folder icon regardless of name', () => {
    expect(fileTypeIcon({ kind: 'folder', name: 'Photos', mime_type: null })).toBe('layers');
  });

  it('maps a plain text document to the document icon', () => {
    expect(
      fileTypeIcon({ kind: 'document', name: 'notes.txt', mime_type: 'text/plain' }),
    ).toBe('document');
  });

  it('maps a spreadsheet MIME type to the spreadsheet icon', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'budget.xlsx',
        mime_type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    ).toBe('bar-chart');
  });

  it('maps a spreadsheet by extension when the MIME type is generic', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'budget.csv',
        mime_type: 'application/octet-stream',
      }),
    ).toBe('bar-chart');
  });

  it('maps a PDF to the PDF icon', () => {
    expect(
      fileTypeIcon({ kind: 'document', name: 'contract.pdf', mime_type: 'application/pdf' }),
    ).toBe('invoice');
  });

  it('maps an image MIME type to the image icon', () => {
    expect(
      fileTypeIcon({ kind: 'document', name: 'photo.png', mime_type: 'image/png' }),
    ).toBe('image');
  });

  it('maps an archive to the archive icon', () => {
    expect(
      fileTypeIcon({ kind: 'document', name: 'backup.zip', mime_type: 'application/zip' }),
    ).toBe('bag');
  });

  it('maps audio by extension to the audio icon', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'voice-memo.mp3',
        mime_type: 'application/octet-stream',
      }),
    ).toBe('bell');
  });

  it('maps video by extension to the video icon', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'clip.mp4',
        mime_type: 'application/octet-stream',
      }),
    ).toBe('layout');
  });

  it('maps source code by extension to the code icon', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'index.ts',
        mime_type: 'text/plain',
      }),
    ).toBe('git-branch');
  });

  it('falls back to the generic document icon for an unrecognised type', () => {
    expect(
      fileTypeIcon({
        kind: 'document',
        name: 'mystery.bin',
        mime_type: 'application/octet-stream',
      }),
    ).toBe('document');
  });

  it('falls back to the generic document icon when mime_type is absent', () => {
    expect(fileTypeIcon({ kind: 'document', name: 'no-extension', mime_type: null })).toBe(
      'document',
    );
  });

  it('every icon this helper can return resolves in the fe-core icon registry', () => {
    const candidates = [
      { kind: 'folder' as const, name: 'x', mime_type: null },
      { kind: 'document' as const, name: 'a.txt', mime_type: 'text/plain' },
      { kind: 'document' as const, name: 'a.xlsx', mime_type: null },
      { kind: 'document' as const, name: 'a.pdf', mime_type: 'application/pdf' },
      { kind: 'document' as const, name: 'a.png', mime_type: 'image/png' },
      { kind: 'document' as const, name: 'a.zip', mime_type: 'application/zip' },
      { kind: 'document' as const, name: 'a.mp3', mime_type: null },
      { kind: 'document' as const, name: 'a.mp4', mime_type: null },
      { kind: 'document' as const, name: 'a.py', mime_type: null },
      { kind: 'document' as const, name: 'a.unknown', mime_type: null },
    ];

    for (const node of candidates) {
      expect(hasIcon(fileTypeIcon(node))).toBe(true);
    }
  });
});

describe('fileTypeIcon — extension fallback for images', () => {
  it('uses the image icon for an image EXTENSION even when the sniffed mime is not an image', () => {
    // The server sniffs the stored mime from the bytes, so a .png whose content
    // is not recognisably an image arrives as text/plain. The name is still
    // unambiguous, so the icon should not degrade to the generic document.
    for (const name of ['logo.png', 'photo.JPG', 'icon.svg', 'anim.gif', 'shot.webp']) {
      expect(fileTypeIcon({ kind: 'document', name, mime_type: 'text/plain' })).toBe('image');
    }
  });

  it('still prefers a real image mime when one is present', () => {
    expect(fileTypeIcon({ kind: 'document', name: 'no-extension', mime_type: 'image/png' })).toBe(
      'image',
    );
  });
});
