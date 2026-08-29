// Formula-editing actions: fast fill, the IFERROR guard, unit scaling, sign
// flip, decimal stepping, CAGR insertion and the consistent-rounding block.
// Shares the selection cap and undo capture with selection.ts rather than
// duplicating either.

import { cappedAreas } from "./areas";
import {
  numberFormat,
  requireEmptyBlock,
  selectedSingleRange,
  SHEET_COLUMNS,
  SHEET_ROWS,
  syncWrite,
} from "./internal";
import { parseAddress } from "./shared";
import { captureUndo, captureUndoAreas } from "./undo";
import { type CellValue, scaleCells } from "../model";
import {
  absoluteRef,
  buildCagrFormula,
  buildRoundFormula,
  detectFillExtent,
  flipSign,
  formatDecimals,
  stepDecimals,
  toggleIfError,
} from "../paste";
import { ROUNDING_CELL_CAP } from "../rounding";

const FILL_SCAN_LIMIT = 1_000;

// What office.js takes back: a grid of literals, never null.
type WritableGrid = (string | number | boolean)[][];

// One grid transform over every area of the selection: read, capture undo, write
// back. A ctrl-clicked pair of blocks is one job, not two.
async function editAreas(
  what: string,
  property: "formulas" | "numberFormat",
  edit: (grid: CellValue[][]) => CellValue[][],
): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await cappedAreas(context, what);
    for (const area of areas) area.load(property);
    await context.sync();
    await captureUndoAreas(context, areas);

    for (const area of areas) {
      const next = edit(area[property] as CellValue[][]) as WritableGrid;
      if (property === "formulas") area.formulas = next;
      else area.numberFormat = next;
    }
    await syncWrite(context, what);
  });
}

// Macabacus-style fast fill: the data beside the origin decides how far the
// formula travels, so nobody has to select the block first.
export async function fastFillAuto(direction: "right" | "down"): Promise<void> {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    const sheet = cell.worksheet;
    cell.load("rowIndex,columnIndex,formulas");
    await context.sync();

    const formula = (cell.formulas as CellValue[][])[0]?.[0] ?? null;
    if (typeof formula !== "string" || !formula.startsWith("=")) {
      throw new Error("The active cell must contain a formula.");
    }

    const { rowIndex, columnIndex } = cell;
    const down = direction === "down";
    const span = Math.min(
      FILL_SCAN_LIMIT,
      down ? SHEET_ROWS - rowIndex : SHEET_COLUMNS - columnIndex,
    );
    const lines = (
      down ? [columnIndex - 1, columnIndex + 1] : [rowIndex - 1, rowIndex + 1]
    )
      .filter(
        (index) => index >= 0 && index < (down ? SHEET_COLUMNS : SHEET_ROWS),
      )
      .map((index) =>
        down
          ? sheet.getRangeByIndexes(rowIndex, index, span, 1)
          : sheet.getRangeByIndexes(index, columnIndex, 1, span),
      );
    for (const line of lines) line.load("values");
    await context.sync();

    const extent = detectFillExtent(
      lines.map((line) => {
        const values = line.values as CellValue[][];
        return down ? values.map((row) => row[0] ?? null) : (values[0] ?? []);
      }),
    );
    if (extent === 0) throw new Error("No neighbor data to size the fill.");

    const destination = down
      ? cell.getResizedRange(extent - 1, 0)
      : cell.getResizedRange(0, extent - 1);
    await captureUndo(context, destination);

    destination.copyFrom(cell, Excel.RangeCopyType.formulas);
    await context.sync();
  });
}

export async function toggleIfErrorGuard(): Promise<void> {
  await editAreas("The IFERROR guard", "formulas", (grid) =>
    toggleIfError(grid, "0"),
  );
}

export async function scaleSelection(factor: 1000 | 0.001): Promise<void> {
  await editAreas("Scaling", "formulas", (grid) => scaleCells(grid, factor));
}

export async function applySignFlip(): Promise<void> {
  await editAreas("Sign flip", "formulas", flipSign);
}

export async function applyDecimalStep(delta: 1 | -1): Promise<void> {
  await editAreas("Decimal stepping", "numberFormat", (grid) =>
    grid.map((row) =>
      row.map((format) => {
        const current = typeof format === "string" ? format : "General";
        // Excel's own Increase Decimal reads General as "0"; stepDecimals, being
        // a pure format transform, leaves an unnumbered format alone.
        const base = current === "General" && delta === 1 ? "0" : current;
        return stepDecimals(base, delta);
      }),
    ),
  );
}

export async function insertCagr(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();

    const { rowCount, columnCount } = range;
    const periods = (rowCount === 1 ? columnCount : rowCount) - 1;
    if ((rowCount !== 1 && columnCount !== 1) || periods < 1) {
      throw new Error("Select one row or column with at least two periods.");
    }

    const first = range.getCell(0, 0);
    const last = range.getCell(rowCount - 1, columnCount - 1);
    first.load("address");
    last.load("address");
    await context.sync();

    // The result lands just past the series, where a growth row usually sits.
    const destination =
      rowCount === 1 ? last.getOffsetRange(0, 1) : last.getOffsetRange(1, 0);
    await captureUndo(context, destination);

    destination.numberFormat = [[numberFormat("percent")]];
    destination.formulas = [
      [
        buildCagrFormula(
          parseAddress(first.address).address,
          parseAddress(last.address).address,
          periods,
        ),
      ],
    ];
    await context.sync();
  });
}

const ROUNDING_SHAPE_ERROR =
  "Consistent rounding: select one row or column with at least two numbers.";

// Every cell of the group carries the whole group as its first argument, so a
// change anywhere in it recalculates all of them; the position is a literal,
// written here rather than read from the cell's own address.
function roundingFormulas(
  reference: string,
  count: number,
  decimals: number,
): string[] {
  return Array.from({ length: count }, (_, index) =>
    buildRoundFormula(reference, index + 1, decimals),
  );
}

// A rounding group is numbers only: a blank or a label in the middle would
// reach the custom function as a zero and quietly join the allocation.
function requireNumbers(range: Excel.Range, count: number): void {
  const cells = (range.values as CellValue[][]).flat();
  if (cells.filter((cell) => typeof cell === "number").length !== count) {
    throw new Error(ROUNDING_SHAPE_ERROR);
  }
}

// think-cell TCROUND in two steps: PLSFIX.ROUND formulas land beside the numbers
// they round, and their results add up to PLSFIX.ROUNDSUM of the same range. The
// precision comes from what the first cell already prints, so the column beside
// a euro or percentage block rounds the way the block reads.
export async function insertConsistentRounding(): Promise<string> {
  return Excel.run(async (context) => {
    const range = await selectedSingleRange(context, "Consistent rounding");
    range.load("address,rowCount,columnCount,values,numberFormat");
    await context.sync();

    const { rowCount, columnCount } = range;
    const acrossRow = rowCount === 1;
    const count = acrossRow ? columnCount : rowCount;
    if ((!acrossRow && columnCount !== 1) || count < 2) {
      throw new Error(ROUNDING_SHAPE_ERROR);
    }
    if (count > ROUNDING_CELL_CAP) {
      throw new Error(
        `Consistent rounding groups up to ${ROUNDING_CELL_CAP} cells at once.`,
      );
    }
    requireNumbers(range, count);

    const format = (range.numberFormat as CellValue[][])[0]?.[0];
    const decimals =
      (typeof format === "string" ? formatDecimals(format) : null) ?? 0;
    const reference = absoluteRef(parseAddress(range.address).address);
    const destination = acrossRow
      ? range.getOffsetRange(1, 0)
      : range.getOffsetRange(0, 1);
    await requireEmptyBlock(
      context,
      destination,
      `Consistent rounding: the cells ${acrossRow ? "below" : "right of"} the selection are not empty.`,
    );
    await captureUndo(context, destination);

    const formulas = roundingFormulas(reference, count, decimals);
    destination.formulas = acrossRow ? [formulas] : formulas.map((f) => [f]);
    await context.sync();

    return `Consistent rounding: ${count} cells at ${decimals} decimals`;
  });
}
