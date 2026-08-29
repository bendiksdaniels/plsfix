// A1 range geometry, pure: an address (one area or several, absolute or not,
// whole rows and whole columns included) read into rectangles, and whether two
// addresses overlap. Auto-push decides with this which links an edit touched.
// No Office.js and no workbook here - addresses arrive already local, without
// their sheet.

export interface Rect {
  // 1-based and inclusive, the way A1 counts.
  top: number;
  left: number;
  bottom: number;
  right: number;
}

// Excel's own grid, so "A:A" and "3:3" have a size to be intersected against.
const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;
const LETTER_BASE = 26;
const LETTER_OFFSET = "A".charCodeAt(0) - 1;
// Both halves optional: "B4", "B" (from "A:C") and "4" (from "3:3") all arrive
// here, and a corner with neither is rejected by the caller.
const CORNER = /^\$?([A-Za-z]{1,3})?\$?([0-9]{1,7})?$/;

interface Corner {
  column: number | null;
  row: number | null;
}

function columnNumber(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * LETTER_BASE + (letter.charCodeAt(0) - LETTER_OFFSET);
  }
  return value;
}

function corner(text: string): Corner | null {
  const match = CORNER.exec(text.trim());
  if (!match) return null;
  const [, letters, digits] = match;
  if (letters === undefined && digits === undefined) return null;
  return {
    column: letters === undefined ? null : columnNumber(letters),
    row: digits === undefined ? null : Number(digits),
  };
}

function box(rows: [number, number], columns: [number, number]): Rect {
  return {
    top: Math.min(...rows),
    bottom: Math.max(...rows),
    left: Math.min(...columns),
    right: Math.max(...columns),
  };
}

// The two corners have to describe the same kind of thing: two cells, two whole
// columns or two whole rows. "A1:B" is not an address we understand, and saying
// so is better than guessing at half of it.
function rectOf(from: Corner, to: Corner): Rect | null {
  const { row: fromRow, column: fromColumn } = from;
  const { row: toRow, column: toColumn } = to;
  const rows = fromRow !== null && toRow !== null;
  const columns = fromColumn !== null && toColumn !== null;
  if (rows && columns) {
    return box([fromRow, toRow], [fromColumn, toColumn]);
  }
  if (rows && fromColumn === null && toColumn === null) {
    return box([fromRow, toRow], [1, MAX_COLUMN]);
  }
  if (columns && fromRow === null && toRow === null) {
    return box([1, MAX_ROW], [fromColumn, toColumn]);
  }
  return null;
}

// Null for anything that is not a single "A1", "A1:B2", "A:C" or "3:5".
export function parseRect(area: string): Rect | null {
  const parts = area.split(":");
  if (parts.length > 2) return null;
  const from = corner(parts[0] ?? "");
  const to = parts.length === 2 ? corner(parts[1] ?? "") : from;
  if (from === null || to === null) return null;
  // A lone corner must be a cell: a bare "B" or "4" names no range.
  if (parts.length === 1 && (from.row === null || from.column === null)) {
    return null;
  }
  return rectOf(from, to);
}

// Null means "this address is not one we can read", which the caller must not
// silently treat as "no overlap".
export function parseAreas(address: string): Rect[] | null {
  const areas: Rect[] = [];
  for (const part of address.split(",")) {
    const rect = parseRect(part);
    if (rect === null) return null;
    areas.push(rect);
  }
  return areas;
}

export function overlaps(a: Rect, b: Rect): boolean {
  return (
    a.left <= b.right &&
    b.left <= a.right &&
    a.top <= b.bottom &&
    b.top <= a.bottom
  );
}

// An address neither side could read counts as a hit: one push too many costs a
// render, one push too few leaves a stale picture in somebody's deck.
export function intersects(anchor: string, changed: string): boolean {
  const left = parseAreas(anchor);
  const right = parseAreas(changed);
  if (left === null || right === null) return true;
  return left.some((a) => right.some((b) => overlaps(a, b)));
}
