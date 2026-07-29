import { describe, it, expect } from 'vitest';
import {
  addressesInRange,
  findContainingRange,
  normalizeRange,
  parseRangeText,
  rangeContainsCell,
  rangeReferenceFor,
} from '../src/utils/sheetAddress';

describe('normalizeRange', () => {
  it('normalises an out-of-order drag into top-left/bottom-right corners', () => {
    expect(normalizeRange({ column: 4, row: 10 }, { column: 2, row: 3 })).toEqual({
      startColumn: 2,
      endColumn: 4,
      startRow: 3,
      endRow: 10,
    });
  });
});

describe('rangeReferenceFor', () => {
  it('a click (same start/end cell) produces a bare address', () => {
    expect(rangeReferenceFor({ column: 2, row: 2 }, { column: 2, row: 2 })).toBe('B2');
  });

  it('a drag produces a normalised A1:C3-style range', () => {
    expect(rangeReferenceFor({ column: 2, row: 10 }, { column: 2, row: 2 })).toBe('B2:B10');
  });
});

describe('parseRangeText', () => {
  it('parses a bare address as a one-cell range', () => {
    expect(parseRangeText('B2')).toEqual({ startColumn: 2, endColumn: 2, startRow: 2, endRow: 2 });
  });

  it('parses a two-corner range regardless of corner order', () => {
    expect(parseRangeText('C10:A1')).toEqual({ startColumn: 1, endColumn: 3, startRow: 1, endRow: 10 });
  });

  it('returns null for a malformed range', () => {
    expect(parseRangeText('not-a-range')).toBeNull();
  });
});

describe('rangeContainsCell / findContainingRange', () => {
  it('rangeContainsCell is true only inside the rectangle', () => {
    const range = { startColumn: 1, endColumn: 3, startRow: 1, endRow: 1 };
    expect(rangeContainsCell(range, 1, 2)).toBe(true);
    expect(rangeContainsCell(range, 2, 4)).toBe(false);
  });

  it('findContainingRange resolves a merged-away cell to its merge range', () => {
    const merges = ['A1:C1', 'E5:E6'];
    expect(findContainingRange(merges, 1, 2)).toEqual({
      startColumn: 1,
      endColumn: 3,
      startRow: 1,
      endRow: 1,
    });
    expect(findContainingRange(merges, 6, 5)).toEqual({
      startColumn: 5,
      endColumn: 5,
      startRow: 5,
      endRow: 6,
    });
  });

  it('findContainingRange returns null when no merge covers the cell', () => {
    expect(findContainingRange(['A1:C1'], 5, 5)).toBeNull();
  });
});

describe('addressesInRange', () => {
  it('lists every address row-major', () => {
    expect(addressesInRange({ startColumn: 1, endColumn: 2, startRow: 1, endRow: 2 })).toEqual([
      'A1',
      'B1',
      'A2',
      'B2',
    ]);
  });
});
