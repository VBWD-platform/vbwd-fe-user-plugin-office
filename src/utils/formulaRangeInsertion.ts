// Formula-range drag insertion (S147-4 drag/toolbar slice, headline feature):
// while the formula bar or an inline cell editor holds text starting with
// "=", dragging over the grid inserts (and, on every further mousemove tick,
// REPLACES) an A1 reference at the cursor — `=SUM(` + drag B2->B10 becomes
// `=SUM(B2:B10)`. Pure text-splicing, no DOM/cursor API here so this is
// trivially unit-testable; `OfficeSheetEditor.vue` owns reading/writing the
// actual `<input>`'s `selectionStart`/value.

/** Tracks one in-progress drag-to-pick-a-range gesture against a single
 * formula string: `anchorPosition` is where the reference was first
 * inserted (the cursor position at drag-start, fixed for the whole drag);
 * `insertedLength` is how many characters of the CURRENT reference text sit
 * at that position, so the next tick knows what to replace. */
export interface FormulaRangeDragState {
  anchorPosition: number;
  insertedLength: number;
}

export function beginRangeInsertion(cursorPosition: number): FormulaRangeDragState {
  return { anchorPosition: cursorPosition, insertedLength: 0 };
}

export interface RangeInsertionResult {
  text: string;
  cursorPosition: number;
  state: FormulaRangeDragState;
}

/** Splice `referenceText` into `text` at `state.anchorPosition`, replacing
 * whatever reference the PREVIOUS tick of the same drag left there
 * (`state.insertedLength`, `0` on the first tick). Returns the new text, the
 * cursor position immediately after the inserted reference, and the updated
 * state to pass into the next tick. */
export function applyRangeInsertion(
  text: string,
  state: FormulaRangeDragState,
  referenceText: string,
): RangeInsertionResult {
  const before = text.slice(0, state.anchorPosition);
  const after = text.slice(state.anchorPosition + state.insertedLength);
  const nextText = `${before}${referenceText}${after}`;
  return {
    text: nextText,
    cursorPosition: state.anchorPosition + referenceText.length,
    state: { anchorPosition: state.anchorPosition, insertedLength: referenceText.length },
  };
}

/** Whether `rawText` is formula input a range-pick drag should attach to —
 * mirrors the store's own `=`-prefix convention (`changeFromRawInput`). */
export function isFormulaInput(rawText: string): boolean {
  return rawText.startsWith('=');
}
