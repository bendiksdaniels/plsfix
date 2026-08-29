// Super Find's matching rules: which cells of a grid and which comments a query
// hits, what each hit reads as, and in what order the workbook's hits are
// listed. Pure: the adapter in src/excel/find.ts reads the sheets and the
// comments and turns these hits into addresses.

import { type CellValue, isFormula } from "./model";

// A results list is read, not scrolled: past a couple of hundred rows the pane
// is a worse tool than Excel's own Find box, and every hit past the cap costs
// another sheet read to produce.
export const FIND_HIT_CAP = 200;

// Names belong to the workbook rather than to any sheet, so they sort ahead of
// every sheet; a sheet-name hit sorts ahead of that sheet's own cells.
export const NAME_ORDER = -1;
export const SHEET_NAME_ROW = -1;
// A sheet-name hit has no cell, so it sorts ahead of column A as well.
export const SHEET_NAME_COL = -1;

export interface FindGrid {
  values: CellValue[][];
  formulas: CellValue[][];
}

export interface MatchOptions {
  matchCase: boolean;
  inFormulas: boolean;
}

export interface CellHit {
  row: number;
  col: number;
  text: string;
}

export interface RankedHit {
  sheetIndex: number;
  row: number;
  col: number;
}

// One rule for cells, defined names and sheet names alike: a substring, matched
// without regard to case unless the modeller asked for it.
export function includesQuery(
  haystack: string,
  query: string,
  matchCase: boolean,
): boolean {
  if (query === "") return false;
  if (matchCase) return haystack.includes(query);
  return haystack.toLowerCase().includes(query.toLowerCase());
}

// What the cell reads as. Booleans are spelled the way Excel shows them, so a
// search for TRUE finds them; numbers are searched as they are stored, not as
// their number format renders them.
export function cellText(value: CellValue): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

// The value is what the modeller sees, so it is searched first and shown as it
// stands. The formula behind it is only searched when the box is ticked, and
// only where the cell really holds one: Excel reports a constant in both grids,
// which would otherwise match the same cell twice.
function matchOne(
  value: CellValue,
  formula: CellValue,
  query: string,
  options: MatchOptions,
): string | null {
  const shown = cellText(value);
  if (includesQuery(shown, query, options.matchCase)) return shown;
  if (!options.inFormulas || !isFormula(formula)) return null;
  return includesQuery(formula, query, options.matchCase) ? formula : null;
}

// Reading order within the sheet, capped: a query that hits everywhere must not
// build a list nothing will ever render.
export function matchCells(
  grid: FindGrid,
  query: string,
  options: MatchOptions,
): CellHit[] {
  const hits: CellHit[] = [];
  if (query === "") return hits;

  for (let row = 0; row < grid.values.length; row += 1) {
    const values = grid.values[row] ?? [];
    const formulas = grid.formulas[row] ?? [];
    for (let col = 0; col < values.length; col += 1) {
      if (hits.length >= FIND_HIT_CAP) return hits;
      const text = matchOne(
        values[col] ?? null,
        formulas[col] ?? null,
        query,
        options,
      );
      if (text !== null) hits.push({ row, col, text });
    }
  }
  return hits;
}

// Workbook reading order: defined names, then each sheet in tab order - its own
// name first, then its cells down and across. The cap lands here too, so the
// merged list is never longer than a single sheet's.
export function rankHits<T extends RankedHit>(hits: T[]): T[] {
  return [...hits]
    .sort(
      (left, right) =>
        left.sheetIndex - right.sheetIndex ||
        left.row - right.row ||
        left.col - right.col,
    )
    .slice(0, FIND_HIT_CAP);
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

// A comment as the workbook hands it over: the cell it hangs on, what it says
// and who said it. A reply is one of these too - the adapter keeps the two
// apart, so a search for "reply" cannot match every reply there is.
export interface CommentEntry {
  sheet: string;
  address: string;
  content: string;
  author: string;
}

export interface CommentHit {
  // Where the comment sat in the list handed in: a comment has no grid
  // coordinates of its own, so this is what carries the workbook's order back
  // to the caller holding the rest of the record.
  order: number;
  text: string;
}

// A thread can run to paragraphs; the row shows its opening, and the modeller
// reads the rest in the cell the hit jumps to.
export const COMMENT_TEXT_LIMIT = 120;

// Comments hang off the grid rather than sit in it, so they sort past Excel's
// last row: a sheet's comments land after its cells, in the workbook's order.
export const COMMENT_ROW = 2_000_000;

// The author is searched beside the content: "who flagged this" is as much a
// question as "what did they say". The cap is the cells' cap, counted here too
// so a workbook papered in comments cannot build a list nothing will render.
export function matchComments(
  comments: CommentEntry[],
  query: string,
  options: MatchOptions,
): CommentHit[] {
  const hits: CommentHit[] = [];
  if (query === "") return hits;

  for (let order = 0; order < comments.length; order += 1) {
    if (hits.length >= FIND_HIT_CAP) return hits;
    const comment = comments[order];
    if (!comment) continue;
    const matched =
      includesQuery(comment.content, query, options.matchCase) ||
      includesQuery(comment.author, query, options.matchCase);
    if (matched) {
      hits.push({ order, text: comment.content.slice(0, COMMENT_TEXT_LIMIT) });
    }
  }
  return hits;
}

const ALPHABET = 26;
const FIRST_LETTER = 65;

// Zero-based grid coordinates to the A1 address the modeller reads, so a hit can
// be shown and jumped to without asking the host for its address.
export function cellAddress(row: number, col: number): string {
  let letters = "";
  for (let rest = col; rest >= 0; rest = Math.floor(rest / ALPHABET) - 1) {
    letters = String.fromCharCode(FIRST_LETTER + (rest % ALPHABET)) + letters;
  }
  return `${letters}${row + 1}`;
}
