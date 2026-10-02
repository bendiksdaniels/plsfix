// What a copied block's formulas become at their new corner. Pure: the block
// and the offset go in, the rewritten formula text comes out.
// Owns the inside/outside rule over src/formula-refs.ts's scanner.
// Invariant: a reference wholly inside the block follows the block, and every
// other reference keeps pointing at the cells it was written for - which on a
// paste onto another sheet means naming the sheet it was written on.

import {
  type FormulaRef,
  formatReference,
  GRID_COLUMNS,
  GRID_ROWS,
  type RefPart,
  rewriteReferences,
  sheetPrefix,
} from "./formula-refs";
import { type CellValue, isFormula } from "./model";

/** A copied rectangle, zero-based, on the sheet it was copied from. */
export interface CellBlock {
  sheet: string;
  row: number;
  column: number;
  rowCount: number;
  columnCount: number;
}

/** How far the paste lands from the block: positive is down and right. */
export interface CellOffset {
  rows: number;
  columns: number;
}

// Excel sheet names do not differ by case, so neither may this test.
function sameSheet(name: string, sheet: string): boolean {
  return sheet !== "" && name.toLowerCase() === sheet.toLowerCase();
}

// A missing half is a whole column ("A:A") or a whole row ("1:1"), which spans
// the sheet - so it is only ever inside a block that spans the sheet too.
function span(
  from: number | null,
  to: number | null,
  size: number,
): [number, number] {
  if (from === null || to === null) return [0, size - 1];
  return [Math.min(from, to), Math.max(from, to)];
}

// Inside means the WHOLE reference is inside: one that straddles the edge
// points partly at cells the paste does not carry, so it must not move.
function insideBlock(ref: FormulaRef, block: CellBlock): boolean {
  const [rowFrom, rowTo] = span(ref.from.row, ref.to.row, GRID_ROWS);
  const [columnFrom, columnTo] = span(
    ref.from.column,
    ref.to.column,
    GRID_COLUMNS,
  );
  return (
    rowFrom >= block.row &&
    rowTo < block.row + block.rowCount &&
    columnFrom >= block.column &&
    columnTo < block.column + block.columnCount
  );
}

function shiftPart(part: RefPart, offset: CellOffset): RefPart | null {
  const row = part.row === null ? null : part.row + offset.rows;
  const column = part.column === null ? null : part.column + offset.columns;
  if (row !== null && (row < 0 || row >= GRID_ROWS)) return null;
  if (column !== null && (column < 0 || column >= GRID_COLUMNS)) return null;
  return { ...part, row, column };
}

// Which sheet each kind of reference has to name once the block lands on
// another sheet. Both are "" for a paste that stays on its own sheet, which is
// the signal to leave every prefix exactly as the modeller wrote it.
interface SheetMove {
  crossed: boolean;
  /** The block's own cells now live here. */
  destination: string;
  /** Everything else still lives here. */
  source: string;
}

// A reference the block owns: unqualified, or qualified with the block's own
// sheet. Anything else already names where it points and is left alone.
function owned(ref: FormulaRef, block: CellBlock): boolean {
  return ref.sheet === "" || sameSheet(ref.sheet, block.sheet);
}

function rewrite(
  ref: FormulaRef,
  block: CellBlock,
  offset: CellOffset,
  move: SheetMove,
): string | null {
  if (!owned(ref, block)) return null;

  if (!insideBlock(ref, block)) {
    // Outside: the address never changes. On another sheet an unqualified one
    // would start reading the destination's cells, so it gains the source
    // sheet; one already qualified with it is right as written.
    if (!move.crossed || ref.prefix !== "") return null;
    return formatReference(ref, ref.from, ref.to, move.source);
  }

  const from = shiftPart(ref.from, offset);
  const to = shiftPart(ref.to, offset);
  // Off the grid: nothing sensible to write, so the reference stays put.
  if (from === null || to === null) return null;
  // Inside: it follows the block. An unqualified one stays unqualified and is
  // already the destination's own cell; one written with the source sheet's
  // name has to be re-pointed at the destination.
  const prefix =
    move.crossed && ref.prefix !== "" ? move.destination : ref.prefix;
  return formatReference(ref, from, to, prefix);
}

/**
 * One formula of a copied block, rewritten for the block's new corner: a
 * reference inside the block moves with it whatever its `$` markers say, and
 * a reference outside it keeps the cells it was written for. A sheet prefix
 * naming another sheet is outside by definition; the block's own sheet by
 * name is inside when the address falls in the block.
 *
 * `toSheet` is the sheet the block is being pasted onto; it defaults to the
 * block's own, which is a paste that stays where it is.
 */
export function duplicateFormula(
  formula: CellValue,
  block: CellBlock,
  offset: CellOffset,
  toSheet: string = block.sheet,
): CellValue {
  if (!isFormula(formula)) return formula;
  // Neither name known means nothing can be said about sheets, so nothing is.
  const crossed =
    block.sheet !== "" && toSheet !== "" && !sameSheet(toSheet, block.sheet);
  const move: SheetMove = {
    crossed,
    destination: crossed ? sheetPrefix(toSheet) : "",
    source: crossed ? sheetPrefix(block.sheet) : "",
  };
  return rewriteReferences(formula, (ref) => rewrite(ref, block, offset, move));
}

export function duplicateFormulas(
  cells: CellValue[][],
  block: CellBlock,
  offset: CellOffset,
  toSheet: string = block.sheet,
): CellValue[][] {
  return cells.map((row) =>
    row.map((cell) => duplicateFormula(cell, block, offset, toSheet)),
  );
}
