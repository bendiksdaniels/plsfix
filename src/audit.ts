// The formula-consistency maths both audit tools share, pure and grid-shaped:
// `auditGrid` classifies every cell of one R1C1 grid for the overlay's stripes,
// `consistentRegion` grows the block "Select consistent region" selects. Two
// cells are the same formula only when their R1C1 strings are identical.

import { type CellValue, isFormula } from "./model";

export type AuditMark =
  "none" | "horizontal" | "vertical" | "both" | "lone" | "typed";

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

// The raw cell at a position, or undefined past this row's own length. That is
// the one thing formulaAt cannot say: it collapses "no such cell" and "a blank
// cell sits there" into the same null, and a typed number's across neighbours
// must tell those apart.
function valueAt(
  grid: CellValue[][],
  row: number,
  column: number,
): CellValue | undefined {
  const line = grid[row];
  if (!line || column < 0 || column >= line.length) return undefined;
  return line[column] ?? null;
}

// The nearest formula across from a cell, stepping over numeric constants -
// an input sitting mid-row does not end the run - but stopping dead at a
// blank or text cell, which is where a modeller's fill really stopped, or at
// the grid edge.
function nearestFormulaAcross(
  grid: CellValue[][],
  row: number,
  column: number,
  step: -1 | 1,
): string | null {
  let index = column + step;
  for (;;) {
    const value = valueAt(grid, row, index);
    if (value === undefined) return null;
    if (typeof value === "number") {
      index += step;
      continue;
    }
    return isFormula(value) ? value : null;
  }
}

// A typed number reads as a hardcode inside a formula row: both across
// neighbours that exist must be formulas, at least one must exist, and where
// both exist they must be the very same R1C1 formula - the row the modeller
// meant to fill straight across. Down is never asked here: an input row
// sitting between two formula rows is an ordinary model shape, not a
// deviation, and lighting it up would be noise.
function isTypedNumber(
  grid: CellValue[][],
  row: number,
  column: number,
): boolean {
  const left = valueAt(grid, row, column - 1);
  const right = valueAt(grid, row, column + 1);
  const leftFormula = left !== undefined && isFormula(left) ? left : null;
  const rightFormula = right !== undefined && isFormula(right) ? right : null;

  if (left !== undefined && leftFormula === null) return false;
  if (right !== undefined && rightFormula === null) return false;
  if (left === undefined && right === undefined) return false;
  if (left !== undefined && right !== undefined) {
    return leftFormula === rightFormula;
  }
  return true;
}

// R1C1 makes a copied formula read identically in every cell it was filled into,
// so equality with a neighbour is the consistency test (UpSlide Formula Audit).
export function auditGrid(formulasR1C1: CellValue[][]): AuditMark[][] {
  return formulasR1C1.map((row, rowIndex) =>
    row.map((cell, columnIndex): AuditMark => {
      if (typeof cell === "number") {
        return isTypedNumber(formulasR1C1, rowIndex, columnIndex)
          ? "typed"
          : "none";
      }
      if (!isFormula(cell)) return "none";

      const across = [
        nearestFormulaAcross(formulasR1C1, rowIndex, columnIndex, -1),
        nearestFormulaAcross(formulasR1C1, rowIndex, columnIndex, 1),
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
    if (formulaAt(grid, row, column) !== formula) return false;
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
    if (formulaAt(grid, row, column) !== formula) return false;
  }
  return true;
}

/**
 * A rectangle of identical R1C1 formulas grown out of one cell, or null when
 * that cell holds no formula at all. The growth is greedy, not a search for
 * the biggest rectangle there is: horizontal first, then vertical, repeated
 * until nothing moves. So where an L of matching cells could become either,
 * the row wins - the same preference auditGrid shows when it calls a cell
 * "horizontal", and the column those cells sit in is never selected.
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
