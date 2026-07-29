import { describe, it, expect } from 'vitest';
import {
  applyRangeInsertion,
  beginRangeInsertion,
  isFormulaInput,
} from '../src/utils/formulaRangeInsertion';

describe('isFormulaInput', () => {
  it('is true only for "="-prefixed text', () => {
    expect(isFormulaInput('=SUM(')).toBe(true);
    expect(isFormulaInput('42')).toBe(false);
  });
});

describe('beginRangeInsertion / applyRangeInsertion', () => {
  it('inserts a reference at the cursor on the first tick', () => {
    const state = beginRangeInsertion(5);
    const result = applyRangeInsertion('=SUM()', state, 'B2');
    expect(result.text).toBe('=SUM(B2)');
    expect(result.cursorPosition).toBe(7);
    expect(result.state).toEqual({ anchorPosition: 5, insertedLength: 2 });
  });

  it('replaces the PREVIOUSLY inserted reference as the drag continues', () => {
    let state = beginRangeInsertion(5);
    let result = applyRangeInsertion('=SUM()', state, 'B2');
    state = result.state;
    result = applyRangeInsertion(result.text, state, 'B2:B10');

    expect(result.text).toBe('=SUM(B2:B10)');
    expect(result.cursorPosition).toBe(11);
  });

  it('preserves text typed after the inserted reference', () => {
    const state = beginRangeInsertion(5);
    const result = applyRangeInsertion('=SUM()+1', state, 'B2:B10');
    expect(result.text).toBe('=SUM(B2:B10)+1');
  });

  it('a fresh insertion after finishing one drag starts a new anchor', () => {
    let state = beginRangeInsertion(5);
    let result = applyRangeInsertion('=SUM()', state, 'B2');
    // Drag finished; user typed a comma then started a second drag at the new cursor.
    const textWithComma = `${result.text.slice(0, 7)},${result.text.slice(7)}`;
    state = beginRangeInsertion(8);
    result = applyRangeInsertion(textWithComma, state, 'C2');
    expect(result.text).toBe('=SUM(B2,C2)');
  });
});
