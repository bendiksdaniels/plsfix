import { type CellValue, isFormula } from "./model";

export type AuditMark = "none" | "horizontal" | "vertical" | "both" | "lone";

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
