// Reshaping a grid between the wide layout a model is read in and the long
// layout a pivot table or a database wants. Pure: the adapter in
// src/excel/reshape.ts only reads the selection and writes the result.

import { type CellValue } from "./model";

const MIN_UNPIVOT_ROWS = 2;
const MIN_UNPIVOT_COLUMNS = 2;

function isBlank(value: CellValue | undefined): boolean {
  return value === undefined || value === null || value === "";
}

// Wide to long: the header row names the columns, the first column names the
// rows, and every populated cell becomes one [row key, column key, value] line
// in reading order. A blank cell states no fact, and a cell whose row or column
// has no key cannot be named, so both are skipped rather than emitted under an
// empty key. Values keep their type: a number stays a number.
export function unpivot(grid: CellValue[][]): CellValue[][] {
  const header = grid[0] ?? [];
  if (grid.length < MIN_UNPIVOT_ROWS || header.length < MIN_UNPIVOT_COLUMNS) {
    throw new Error(
      "unpivot: need a header row, a key column and one column of values",
    );
  }

  const long: CellValue[][] = [];
  for (const row of grid.slice(1)) {
    const rowKey = row[0] ?? null;
    if (isBlank(rowKey)) continue;

    for (let column = 1; column < header.length; column += 1) {
      const columnKey = header[column] ?? null;
      const value = row[column] ?? null;
      if (isBlank(columnKey) || isBlank(value)) continue;
      long.push([rowKey, columnKey, value]);
    }
  }
  return long;
}
