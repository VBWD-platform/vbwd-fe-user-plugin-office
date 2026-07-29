// A1-style cell address helpers for the VBWD Spreadsheets grid (S147-4).
// Mirrors the backend's `office/sheet/cell.py` column<->index conversion —
// deliberately re-implemented here (not shared code across the Python/TS
// boundary) rather than round-tripping through an HTTP call for something
// this cheap and pure.

const COLUMN_LETTERS_PATTERN = /^([A-Za-z]{1,3})([1-9][0-9]*)$/;

export function columnLettersToIndex(letters: string): number {
  let index = 0;
  for (const character of letters.toUpperCase()) {
    index = index * 26 + (character.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
  }
  return index;
}

export function columnIndexToLetters(index: number): string {
  let letters = '';
  let remaining = index;
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    letters = String.fromCharCode('A'.charCodeAt(0) + remainder) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

export interface SheetCellAddress {
  column: number;
  row: number;
}

/** Parse a bare A1 reference (no sheet qualifier — the grid always knows
 * which sheet it is editing). Returns `null` for anything malformed. */
export function parseA1(address: string): SheetCellAddress | null {
  const match = COLUMN_LETTERS_PATTERN.exec(address.trim());
  if (!match) return null;
  return { column: columnLettersToIndex(match[1]), row: parseInt(match[2], 10) };
}

export function formatA1(column: number, row: number): string {
  return `${columnIndexToLetters(column)}${row}`;
}

/** A rectangular block of cells in ROW/COLUMN index space — the drag-select,
 * fill-handle, and merge/unmerge features all reason about a range as a pair
 * of normalised (top-left, bottom-right) corners rather than the two raw
 * drag endpoints, which may arrive in any order. */
export interface SheetCellRange {
  startColumn: number;
  endColumn: number;
  startRow: number;
  endRow: number;
}

/** Normalise two (possibly out-of-order) corners into a top-left/bottom-right
 * range — the single place every drag gesture (select, fill, formula-range
 * pick) turns two endpoints into a rectangle. */
export function normalizeRange(anchor: SheetCellAddress, focus: SheetCellAddress): SheetCellRange {
  return {
    startColumn: Math.min(anchor.column, focus.column),
    endColumn: Math.max(anchor.column, focus.column),
    startRow: Math.min(anchor.row, focus.row),
    endRow: Math.max(anchor.row, focus.row),
  };
}

/** The A1 reference text for a drag from `anchor` to `focus` — a bare
 * address (`"B2"`) when both corners are the same cell (a click, not a
 * drag), otherwise a `"B2:B10"` range, per the formula-range-insertion and
 * fill-handle interactions. */
export function rangeReferenceFor(anchor: SheetCellAddress, focus: SheetCellAddress): string {
  if (anchor.row === focus.row && anchor.column === focus.column) {
    return formatA1(anchor.column, anchor.row);
  }
  const range = normalizeRange(anchor, focus);
  return `${formatA1(range.startColumn, range.startRow)}:${formatA1(range.endColumn, range.endRow)}`;
}

/** Parse a bare or `"A1:C3"` range string (no sheet qualifier) into a
 * normalised range — the mirror-image reader for whatever
 * {@link rangeReferenceFor} / the backend's merge storage produced. Returns
 * `null` for anything malformed rather than throwing, since callers use it
 * on server-supplied merge strings that are always well-formed in practice
 * but should never crash rendering if they are not. */
export function parseRangeText(rangeText: string): SheetCellRange | null {
  const [startText, endText] = rangeText.split(':');
  const start = parseA1(startText);
  const end = parseA1(endText ?? startText);
  if (!start || !end) return null;
  return normalizeRange(start, end);
}

export function rangeContainsCell(range: SheetCellRange, row: number, column: number): boolean {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  );
}

/** The first range in `ranges` (server order) covering `(row, column)`, or
 * `null` — used to resolve a merged-away cell back to the merge it belongs
 * to (S147-4 drag/toolbar slice: merges are non-overlapping by construction
 * on the backend, so "first match" is also the only match). */
export function findContainingRange(
  ranges: string[],
  row: number,
  column: number,
): SheetCellRange | null {
  for (const rangeText of ranges) {
    const range = parseRangeText(rangeText);
    if (range && rangeContainsCell(range, row, column)) return range;
  }
  return null;
}

/** Every A1 address inside a range, row-major — used to expand a toolbar
 * selection (bold/format/etc.) into the individual cell addresses the
 * backend's `PUT .../cells` contract expects. */
export function addressesInRange(range: SheetCellRange): string[] {
  const addresses: string[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      addresses.push(formatA1(column, row));
    }
  }
  return addresses;
}
