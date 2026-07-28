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
