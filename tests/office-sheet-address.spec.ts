import { describe, it, expect } from 'vitest';
import {
  columnIndexToLetters,
  columnLettersToIndex,
  formatA1,
  parseA1,
} from '../src/utils/sheetAddress';

describe('sheetAddress', () => {
  it('columnLettersToIndex converts single and multi-letter columns', () => {
    expect(columnLettersToIndex('A')).toBe(1);
    expect(columnLettersToIndex('Z')).toBe(26);
    expect(columnLettersToIndex('AA')).toBe(27);
  });

  it('columnIndexToLetters is the inverse of columnLettersToIndex', () => {
    expect(columnIndexToLetters(1)).toBe('A');
    expect(columnIndexToLetters(26)).toBe('Z');
    expect(columnIndexToLetters(27)).toBe('AA');
  });

  it('parseA1 parses a bare cell reference', () => {
    expect(parseA1('B12')).toEqual({ column: 2, row: 12 });
  });

  it('parseA1 returns null for a malformed reference', () => {
    expect(parseA1('not-a-cell')).toBeNull();
    expect(parseA1('12A')).toBeNull();
  });

  it('formatA1 renders a column/row pair back to A1 text', () => {
    expect(formatA1(1, 1)).toBe('A1');
    expect(formatA1(27, 3)).toBe('AA3');
  });
});
