// The formula-consistency maths both audit tools share, pure and grid-shaped.
// Owns two readings of one R1C1 grid: `auditGrid` classifies every cell for the
// overlay's stripes, `consistentRegion` grows the rectangle "Select consistent
// region" hands to Excel. Invariant: two cells are the same formula only when
// their R1C1 strings are identical, which is what a filled range gives.

import { type CellValue, isFormula } from "./model";

export type AuditMark = "none" | "horizontal" | "vertical" | "both" | "lone";

/** A rectangle in grid coordinates, shaped for Range.getRangeByIndexes. */
export interface GridRegion {
  row: number;
  column: number;
  rowCount: number;
  columnCount: number;
}

function formulaAt(
  grid: CellValue[][],
  row: number,
  column: number,
): string | null {
  const cell = grid[row]?.[column] ?? null;
  return isFormula(cell) ? cell : null;
}

// R1C1 makes a copied formula read identically in every cell it was filled into,
// so equality with a neighbour is the consistency test (UpSlide Formula Audit).
export function auditGrid(formulasR1C1: CellValue[][]): AuditMark[][] {
  return formulasR1C1.map((row, rowIndex) =>
    row.map((cell, columnIndex): AuditMark => {
      if (!isFormula(cell)) return "none";

      const across = [
        formulaAt(formulasR1C1, rowIndex, columnIndex - 1),
        formulaAt(formulasR1C1, rowIndex, columnIndex + 1),
      ];
      const down = [
        formulaAt(formulasR1C1, rowIndex - 1, columnIndex),
        formulaAt(formulasR1C1, rowIndex + 1, columnIndex),
      ];

      // A formula with no formula neighbour is unremarkable, not a deviation.
      if (![...across, ...down].some((neighbour) => neighbour !== null)) {
        return "none";
      }

      const matchesAcross = across.includes(cell);
      const matchesDown = down.includes(cell);
      if (matchesAcross && matchesDown) return "both";
      if (matchesAcross) return "horizontal";
      if (matchesDown) return "vertical";
      return "lone";
    }),
  );
}

// A whole row or column of the next step, all of it or none: one blank or one
// different formula is where the fill the modeller made actually stopped.
function sameAcross(
  grid: CellValue[][],
  row: number,
  from: number,
  to: number,
  formula: string,
): boolean {
  for (let column = from; column <= to; column += 1) {
    if ((grid[row]?.[column] ?? null) !== formula) return false;
  }
  return true;
}

function sameDown(
  grid: CellValue[][],
  column: number,
  from: number,
  to: number,
  formula: string,
): boolean {
  for (let row = from; row <= to; row += 1) {
    if ((grid[row]?.[column] ?? null) !== formula) return false;
  }
  return true;
}

/**
 * The largest rectangle of identical R1C1 formulas around one cell, or null
 * when that cell holds no formula at all. Growth is horizontal first, then
 * vertical, repeated until the rectangle stops changing: where an L of
 * matching cells could become either, the row wins - the same preference
 * auditGrid shows when it calls a cell "horizontal".
 */
export function consistentRegion(
  formulasR1C1: CellValue[][],
  row: number,
  column: number,
): GridRegion | null {
  const formula = formulasR1C1[row]?.[column] ?? null;
  if (!isFormula(formula)) return null;

  let top = row;
  let bottom = row;
  let left = column;
  let right = column;
  let growing = true;
  while (growing) {
    growing = false;
    if (sameDown(formulasR1C1, left - 1, top, bottom, formula)) {
      left -= 1;
      growing = true;
    }
    if (sameDown(formulasR1C1, right + 1, top, bottom, formula)) {
      right += 1;
      growing = true;
    }
    if (sameAcross(formulasR1C1, top - 1, left, right, formula)) {
      top -= 1;
      growing = true;
    }
    if (sameAcross(formulasR1C1, bottom + 1, left, right, formula)) {
      bottom += 1;
      growing = true;
    }
  }
  return {
    row: top,
    column: left,
    rowCount: bottom - top + 1,
    columnCount: right - left + 1,
  };
}
