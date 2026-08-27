export type CellValue = string | number | boolean | null;

export interface GridSummary {
  cells: number;
  formulas: number;
  errors: number;
  blanks: number;
}

export function isFormula(value: CellValue): value is string {
  return typeof value === "string" && value.startsWith("=");
}

export function analyzeGrid(
  formulas: CellValue[][],
  values: CellValue[][],
): GridSummary {
  let cells = 0;
  let formulaCount = 0;
  let errors = 0;
  let blanks = 0;

  for (let row = 0; row < formulas.length; row += 1) {
    const formulaRow = formulas[row] ?? [];
    const valueRow = values[row] ?? [];

    for (let column = 0; column < formulaRow.length; column += 1) {
      const formula = formulaRow[column] ?? null;
      const value = valueRow[column] ?? null;
      cells += 1;

      if (isFormula(formula)) formulaCount += 1;
      if (value === null || value === "") blanks += 1;
      if (typeof value === "string" && value.startsWith("#")) errors += 1;
    }
  }

  return { cells, formulas: formulaCount, errors, blanks };
}

export function scaleCells(
  cells: CellValue[][],
  factor: 1000 | 0.001,
): CellValue[][] {
  const operator = factor === 1000 ? "*1000" : "/1000";

  return cells.map((row) =>
    row.map((cell) => {
      if (typeof cell === "number") return cell * factor;
      if (isFormula(cell)) return `=(${cell.slice(1)})${operator}`;
      return cell;
    }),
  );
}

export function makeFormatGrid(
  rows: number,
  columns: number,
  format: string,
): string[][] {
  return Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => format),
  );
}
