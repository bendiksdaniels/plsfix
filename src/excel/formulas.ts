// Formula-editing actions: fast fill, the IFERROR guard, unit scaling, sign flip,
// decimal stepping and CAGR insertion. Shares the selection cap and undo capture
// with selection.ts rather than duplicating either.

import { numberFormat, selectionWithinCap } from "./internal";
import { parseAddress } from "./shared";
import { captureUndo } from "./undo";
import { type CellValue, scaleCells } from "../model";
import {
  buildCagrFormula,
  detectFillExtent,
  flipSign,
  stepDecimals,
  toggleIfError,
} from "../paste";

const FILL_SCAN_LIMIT = 1_000;
const SHEET_ROWS = 1_048_576;
const SHEET_COLUMNS = 16_384;

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
    const lines = (down ? [columnIndex - 1, columnIndex + 1] : [rowIndex - 1, rowIndex + 1])
      .filter((index) => index >= 0 && index < (down ? SHEET_COLUMNS : SHEET_ROWS))
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
        return down ? values.map((row) => row[0] ?? null) : values[0] ?? [];
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
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "The IFERROR guard");
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = toggleIfError(
      range.formulas as CellValue[][],
      "0",
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function scaleSelection(factor: 1000 | 0.001): Promise<void> {
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "Scaling");
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = scaleCells(
      range.formulas as CellValue[][],
      factor,
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function applySignFlip(): Promise<void> {
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "Sign flip");
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = flipSign(
      range.formulas as CellValue[][],
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function applyDecimalStep(delta: 1 | -1): Promise<void> {
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "Decimal stepping");
    range.load("numberFormat");
    await context.sync();
    await captureUndo(context, range);

    range.numberFormat = (range.numberFormat as CellValue[][]).map((row) =>
      row.map((format) => {
        const current = typeof format === "string" ? format : "General";
        // Excel's own Increase Decimal reads General as "0"; stepDecimals, being
        // a pure format transform, leaves an unnumbered format alone.
        const base = current === "General" && delta === 1 ? "0" : current;
        return stepDecimals(base, delta);
      }),
    );
    await context.sync();
  });
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
