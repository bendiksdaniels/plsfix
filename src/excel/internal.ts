// Internal API of the src/excel/ folder: private range, fill and host-capability
// helpers no pane code calls directly. Exported so sibling section files (and
// src/excel/links.ts later) can import them - the barrel never re-exports this
// module, so nothing here is part of the pane's public surface.

import { currencyNumberFormat, getActiveSettings } from "../settings";
import { type NumberFormatName } from "./shared";

const staticNumberFormats = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  percent: "0.0%;[Red](0.0%);-",
} as const;

export function numberFormat(name: NumberFormatName): string {
  if (name === "currency") {
    return currencyNumberFormat(getActiveSettings().currency);
  }
  return staticNumberFormats[name];
}

export const SELECTION_CELL_CAP = 5_000;
export const EDIT_CELL_CAP = 500;
const NO_FILL = "none";
export const BASE_WHITE = "#FFFFFF";

// A whole-column click selects a million cells; reading or writing their grids
// would freeze the pane or overflow the request payload.
export async function selectionWithinCap(
  context: Excel.RequestContext,
  what: string,
): Promise<Excel.Range> {
  const range = context.workbook.getSelectedRange();
  range.load("cellCount");
  await context.sync();
  if (range.cellCount > SELECTION_CELL_CAP) {
    throw new Error(
      `${what} supports up to ${SELECTION_CELL_CAP.toLocaleString()} selected cells at once.`,
    );
  }
  return range;
}

// One write per run of same-key cells instead of one per cell: model rows are
// usually uniform, so this keeps the batch small on wide selections. A null key
// leaves the cell untouched.
export function writeRuns(
  range: Excel.Range,
  keys: (string | null)[][],
  write: (block: Excel.Range, key: string) => void,
): void {
  keys.forEach((row, rowIndex) => {
    let start = 0;
    while (start < row.length) {
      const key = row[start] ?? null;
      let end = start + 1;
      while (end < row.length && (row[end] ?? null) === key) end += 1;
      if (key !== null) {
        write(
          range.getCell(rowIndex, start).getResizedRange(0, end - start - 1),
          key,
        );
      }
      start = end;
    }
  });
}

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was.
export function fillKey(fill: Excel.CellPropertiesFill | undefined): string {
  const pattern = fill?.pattern ?? Excel.FillPattern.none;
  if (pattern === Excel.FillPattern.none) return NO_FILL;
  return [
    pattern,
    fill?.color ?? BASE_WHITE,
    fill?.patternColor ?? BASE_WHITE,
  ].join("|");
}

export function applyFillKey(block: Excel.Range, key: string): void {
  const { fill } = block.format;
  if (key === NO_FILL) {
    fill.clear();
    return;
  }
  const [pattern, color, patternColor] = key.split("|");
  // Colour first: setting it on an unfilled cell would otherwise force Solid.
  fill.color = color ?? BASE_WHITE;
  fill.pattern = pattern as Excel.FillPattern;
  fill.patternColor = patternColor ?? BASE_WHITE;
}

export function hostSupports(apiSet: string): boolean {
  const requirements = Office.context?.requirements;
  return requirements ? requirements.isSetSupported("ExcelApi", apiSet) : true;
}
