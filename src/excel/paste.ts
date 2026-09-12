// Copy/paste actions: marking a copy source, paste special (values/formats/
// transpose), formula-preserving paste, and the three narrow pastes - the
// block's formulas duplicated to a new corner, its number formats only, its
// row heights only. The source is remembered by sheet id and address, not a
// live Excel reference, so it survives a sheet rename.

import { cappedAreas, selectedAreas } from "./areas";
import { withinCap } from "./internal";
import { syncWrite } from "./protection";
import { parseAddress } from "./shared";
import { captureUndoAreas } from "./undo";
import { type CellBlock, duplicateFormulas, tileGrid } from "../paste";
import type { CellValue } from "../model";

export type PasteMode = "values" | "formats" | "transpose";

const PASTE = "Paste";
// A row height is one property read and one band write per row, so this run is
// capped by rows rather than by the cell count every grid flow counts.
const ROW_HEIGHT_ROW_CAP = 500;
const OVERLAP_REFUSAL = "Paste target overlaps the copied block.";

interface CopySource {
  sheetId: string;
  address: string;
  label: string;
}

let copySource: CopySource | null = null;

export function copySourceLabel(): string | null {
  return copySource?.label ?? null;
}

export async function markCopySource(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    sheet.load("id");
    range.load("address");
    await context.sync();

    copySource = {
      sheetId: sheet.id,
      address: parseAddress(range.address).address,
      label: range.address,
    };
    return range.address;
  });
}

// Resolved by sheet id, so renaming the source sheet between copy and paste is fine.
async function openCopySource(
  context: Excel.RequestContext,
): Promise<Excel.Range> {
  const source = copySource;
  if (!source) throw new Error("Mark a copy source first.");

  const sheet = context.workbook.worksheets.getItemOrNullObject(source.sheetId);
  sheet.load("isNullObject");
  await context.sync();

  if (sheet.isNullObject) {
    copySource = null;
    throw new Error("The copy source sheet is gone. Mark a new source.");
  }
  return sheet.getRange(source.address);
}

// Read inside the call, never at module scope: office.js defines the enums, and
// the pane also loads in a plain browser where they do not exist yet.
function pasteCopyType(mode: PasteMode): Excel.RangeCopyType {
  switch (mode) {
    case "values":
      return Excel.RangeCopyType.values;
    case "formats":
      return Excel.RangeCopyType.formats;
    case "transpose":
      return Excel.RangeCopyType.all;
  }
}

export async function pasteSpecial(mode: PasteMode): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    from.load("rowCount,columnCount");
    for (const target of targets) target.load("rowCount,columnCount");
    await context.sync();

    // Excel grows a smaller destination to the source shape, so undo has to
    // cover the whole footprint, not just what the user selected - once per
    // area, because every area of the selection takes a copy.
    const transposed = mode === "transpose";
    const rows = transposed ? from.columnCount : from.rowCount;
    const columns = transposed ? from.rowCount : from.columnCount;
    await captureUndoAreas(
      context,
      targets.map((target) =>
        target
          .getCell(0, 0)
          .getResizedRange(
            Math.max(rows, target.rowCount) - 1,
            Math.max(columns, target.columnCount) - 1,
          ),
      ),
    );

    for (const target of targets) {
      target.copyFrom(from, pasteCopyType(mode), false, transposed);
    }
    await syncWrite(context, PASTE);
  });
}

// Excel's own paste rewrites relative references; this one keeps the formula
// text byte for byte, which is what a modeller means by "same formula here".
export async function pastePreserveFormulas(): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    // Unlike copyFrom, this one carries the source's formulas through the pane,
    // so the source is capped the way every other grid read is. Marking a whole
    // column stays fine: the three copyFrom pastes never leave the host.
    await withinCap(context, from, PASTE);
    from.load("rowCount,columnCount,formulas");
    await context.sync();

    const destinations = targets.map((target) =>
      target
        .getCell(0, 0)
        .getResizedRange(from.rowCount - 1, from.columnCount - 1),
    );
    await captureUndoAreas(context, destinations);

    for (const destination of destinations)
      destination.formulas = from.formulas;
    await syncWrite(context, PASTE);
  });
}

// The destination has the source's shape, so two rectangles of one size, on
// one sheet: a paste that touched the block would read cells it had already
// overwritten and rewrite their references twice.
function overlapsBlock(block: CellBlock, row: number, column: number): boolean {
  return (
    row < block.row + block.rowCount &&
    block.row < row + block.rowCount &&
    column < block.column + block.columnCount &&
    block.column < column + block.columnCount
  );
}

/**
 * The copied block's formulas at the selection's corner, with every reference
 * pointing inside the block moved with it and every reference pointing outside
 * it left at the cells it was written for (src/paste.ts duplicateFormula).
 */
export async function pasteDuplicateFormulas(): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    // The formulas travel through the pane, so the source is capped the way
    // every other grid read is.
    await withinCap(context, from, PASTE);
    const sourceSheet = from.worksheet;
    const targetSheets = targets.map((target) => target.worksheet);
    sourceSheet.load("id,name");
    for (const sheet of targetSheets) sheet.load("id");
    from.load("rowIndex,columnIndex,rowCount,columnCount,formulas");
    for (const target of targets) target.load("rowIndex,columnIndex");
    await context.sync();

    const block: CellBlock = {
      sheet: sourceSheet.name,
      row: from.rowIndex,
      column: from.columnIndex,
      rowCount: from.rowCount,
      columnCount: from.columnCount,
    };
    const formulas = from.formulas as CellValue[][];
    const moves = targets.map((target, index) => ({
      onSourceSheet: targetSheets[index]?.id === sourceSheet.id,
      row: target.rowIndex,
      column: target.columnIndex,
      destination: target
        .getCell(0, 0)
        .getResizedRange(block.rowCount - 1, block.columnCount - 1),
    }));
    for (const move of moves) {
      if (move.onSourceSheet && overlapsBlock(block, move.row, move.column)) {
        throw new Error(OVERLAP_REFUSAL);
      }
    }

    await captureUndoAreas(
      context,
      moves.map((move) => move.destination),
    );
    for (const move of moves) {
      move.destination.formulas = duplicateFormulas(formulas, block, {
        rows: move.row - block.row,
        columns: move.column - block.column,
      });
    }
    await syncWrite(context, PASTE);
  });
}

/**
 * The copied block's number formats only: values, formulas and fills stay as
 * they are. The grid tiles over a selection larger than the source, the way
 * Excel's own copy does, and grows a smaller one to the source shape.
 */
export async function pasteNumberFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    // Both grids travel through the pane here, unlike the copyFrom pastes, so
    // the selection is capped as well as the source.
    const targets = await cappedAreas(context, PASTE);
    await withinCap(context, from, PASTE);
    from.load("rowCount,columnCount,numberFormat");
    for (const target of targets) target.load("rowCount,columnCount");
    await context.sync();

    const formats = from.numberFormat as string[][];
    const shapes = targets.map((target) => ({
      rows: Math.max(from.rowCount, target.rowCount),
      columns: Math.max(from.columnCount, target.columnCount),
    }));
    const destinations = targets.map((target, index) =>
      target
        .getCell(0, 0)
        .getResizedRange(
          (shapes[index]?.rows ?? 1) - 1,
          (shapes[index]?.columns ?? 1) - 1,
        ),
    );
    await captureUndoAreas(context, destinations);

    destinations.forEach((destination, index) => {
      const shape = shapes[index];
      if (!shape) return;
      destination.numberFormat = tileGrid(formats, shape.rows, shape.columns);
    });
    await syncWrite(context, PASTE);
  });
}

function plural(count: number, one: string): string {
  return `${String(count)} ${one}${count === 1 ? "" : "s"}`;
}

function rowHeightReport(written: number, skipped: number): string {
  const skippedNote =
    skipped === 0 ? "" : `, ${plural(skipped, "hidden row")} skipped`;
  // Row heights are sheet state, which getCellProperties cannot carry, so this
  // action runs outside the pls,fix undo net (src/excel/sizes.ts) and says so.
  return `${PASTE}: ${plural(written, "row height")}${skippedNote} (outside pls,fix Undo)`;
}

/**
 * The copied block's row heights onto the rows the selection starts at, one
 * row at a time: a height written on a whole band sets every row in it to the
 * same value (lessons 27.08). A hidden source row carries no height a
 * modeller meant to copy, so it is skipped and counted.
 *
 * Outside pls,fix Undo by design, exactly like src/excel/sizes.ts: row heights
 * are sheet state and captureUndo cannot put one back.
 */
export async function pasteRowHeights(): Promise<string> {
  return Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    from.load("rowCount");
    await context.sync();

    if (from.rowCount > ROW_HEIGHT_ROW_CAP) {
      throw new Error(
        `${PASTE} supports up to ${ROW_HEIGHT_ROW_CAP.toLocaleString()} rows of heights at once.`,
      );
    }

    const rows = Array.from({ length: from.rowCount }, (_unused, index) =>
      from.getRow(index),
    );
    for (const row of rows) row.load("format/rowHeight,rowHidden");
    await context.sync();

    let written = 0;
    let skipped = 0;
    rows.forEach((row, index) => {
      if (row.rowHidden) {
        skipped += 1;
        return;
      }
      const height = row.format.rowHeight;
      for (const target of targets) {
        target
          .getCell(0, 0)
          .getOffsetRange(index, 0)
          .getEntireRow().format.rowHeight = height;
      }
      written += 1;
    });
    await syncWrite(context, PASTE);
    return rowHeightReport(written, skipped);
  });
}
