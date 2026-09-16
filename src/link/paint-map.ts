// Which cells of a table link pls,fix filled with a colour, packed small
// enough for one tag value: a repaint reads this back to know what it is
// allowed to clear, so a fill the deck's table style is showing - never
// written by pls,fix - survives every update. Pure codec, no Office.js.

import { TAG_VALUE_MAX, type TableCell } from "./model";

const PAINT_VERSION = "1";

export function paintKey(row: number, column: number): string {
  return `${String(row)},${String(column)}`;
}

// Consecutive filled columns fold into one "start-end" token, so a filled
// header row costs a few characters rather than one token per cell. Never
// called with an empty list: encodePaintMap only invokes it once it has
// found at least one filled column.
function encodeRanges(columns: number[]): string {
  const [first, ...rest] = columns;
  if (first === undefined) return "";
  const tokens: string[] = [];
  let start = first;
  let end = first;
  const flush = (): void => {
    tokens.push(
      start === end ? `${String(start)}` : `${String(start)}-${String(end)}`,
    );
  };
  for (const column of rest) {
    if (column === end + 1) {
      end = column;
      continue;
    }
    flush();
    start = column;
    end = column;
  }
  flush();
  return tokens.join(",");
}

// "1;" plus one "row:ranges" token per row that carries a fill, ascending -
// "1;" alone is a grid with no filled cell at all.
export function encodePaintMap(cells: TableCell[][]): string {
  const rows: string[] = [];
  cells.forEach((row, r) => {
    const filled = row
      .map((cell, c) => (cell.f !== undefined ? c : -1))
      .filter((c) => c >= 0);
    if (filled.length > 0) rows.push(`${String(r)}:${encodeRanges(filled)}`);
  });
  const encoded = `${PAINT_VERSION};${rows.join(";")}`;
  if (encoded.length > TAG_VALUE_MAX) {
    throw new Error(
      `encodePaintMap: value exceeds ${String(TAG_VALUE_MAX)} chars`,
    );
  }
  return encoded;
}

function addRange(keys: Set<string>, row: number, range: string): boolean {
  const [startText, endText] = range.split("-");
  const start = Number(startText);
  const end = endText === undefined ? start : Number(endText);
  if (!Number.isInteger(start) || !Number.isInteger(end)) return false;
  if (start < 0 || end < start) return false;
  for (let column = start; column <= end; column += 1) {
    keys.add(paintKey(row, column));
  }
  return true;
}

// Absent, malformed or a version this build does not know all read as legacy
// (null): a deck painted before this shipped, or a tag nothing here wrote.
export function decodePaintMap(
  value: string | null | undefined,
): Set<string> | null {
  if (value === null || value === undefined) return null;
  const sep = value.indexOf(";");
  if (sep === -1 || value.slice(0, sep) !== PAINT_VERSION) return null;
  const body = value.slice(sep + 1);
  const keys = new Set<string>();
  if (body === "") return keys;
  for (const part of body.split(";")) {
    const rowSep = part.indexOf(":");
    if (rowSep === -1) return null;
    const row = Number(part.slice(0, rowSep));
    if (!Number.isInteger(row) || row < 0) return null;
    for (const range of part.slice(rowSep + 1).split(",")) {
      if (!addRange(keys, row, range)) return null;
    }
  }
  return keys;
}
