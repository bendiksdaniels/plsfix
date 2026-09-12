// Comps stats: the six statistics a comparables page ends on - min, quartiles,
// median, mean and max - written as LIVE formulas one blank row under the
// selected table, so an edited multiple moves them.
//
// Owns: the Office.js side only (the cap, the refusals, the Undo capture and
// the two write batches). Which columns hold numbers and what each formula
// says is pure, in src/stats.ts. Invariant: the formulas span the data rows of
// the selection, never its header row.

import {
  requireEmptyBlock,
  selectedSingleRange,
  SHEET_ROWS,
  withinCap,
  writeRuns,
} from "./internal";
import { applyPresetFormat, type PresetLook } from "./presets";
import { syncWrite } from "./protection";
import { captureUndo } from "./undo";
import { type CellValue } from "../model";
import {
  type CompsBlock,
  readCompsBlock,
  STATS_GAP_ROWS,
  STATS_ROWS,
  statsGrid,
  statsPresets,
} from "../stats";

const STAGE = "Comps stats";
const GENERAL = "General";
const NOT_EMPTY = "Comps stats need six empty rows under the block.";

// Every statistic of a column is the same kind of number as the column itself,
// so it wears the format of the last data row - the row a modeller formats last
// and the one a total under the block would copy.
function statsFormats(
  formats: string[][],
  block: CompsBlock,
  columnCount: number,
): string[][] {
  const last = formats[formats.length - 1] ?? [];
  const numeric = new Set(block.numericColumns);
  const row = Array.from({ length: columnCount }, (_unused, column) =>
    numeric.has(column) ? (last[column] ?? GENERAL) : GENERAL,
  );
  return Array.from({ length: STATS_ROWS }, () => [...row]);
}

function written(block: CompsBlock): string {
  const columns = block.numericColumns.length;
  const rows = block.lastDataRow - block.firstDataRow + 1;
  const column = columns === 1 ? "column" : "columns";
  return `${STAGE} written: ${columns} ${column} over ${rows} rows`;
}

/**
 * Writes the six statistics rows under the selected comps table and selects
 * them. pls,fix Undo captures whatever stood in those rows first.
 */
export async function insertCompsStats(): Promise<string> {
  return Excel.run(async (context) => {
    // The cap answers before the values are asked for: a clicked column header
    // is a million cells, and the table's own shape only reads after that.
    const range = await withinCap(
      context,
      await selectedSingleRange(context, STAGE),
      STAGE,
    );
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values,numberFormat");
    await context.sync();

    const block = readCompsBlock(
      range.values as CellValue[][],
      range.rowIndex,
      range.columnIndex,
    );
    const top = range.rowIndex + range.rowCount + STATS_GAP_ROWS;
    if (top + STATS_ROWS > SHEET_ROWS) {
      throw new Error(`${STAGE}: no room under the selection`);
    }

    const target = sheet.getRangeByIndexes(
      top,
      range.columnIndex,
      STATS_ROWS,
      range.columnCount,
    );
    await requireEmptyBlock(context, target, NOT_EMPTY);
    await captureUndo(context, target);

    target.formulas = statsGrid(block, range.columnCount);
    // A locked sheet refuses the formulas with a host string that names neither
    // the sheet nor the way out; the flow says both itself.
    await syncWrite(context, STAGE);

    target.numberFormat = statsFormats(
      range.numberFormat as string[][],
      block,
      range.columnCount,
    );
    // A null look leaves the cell alone: the columns that hold no numbers keep
    // whatever formatting was already under them.
    writeRuns<PresetLook>(
      target,
      statsPresets(block, range.columnCount),
      (cells, look) => {
        applyPresetFormat(cells.format, look);
      },
    );
    target.select();
    await syncWrite(context, STAGE);

    return written(block);
  });
}
