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
import { protectedNote, sheetProtected, syncWrite } from "./protection";
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
// and the one a total under the block would copy. A column that holds no
// statistic answers null and is never written to, so the label column and the
// text columns keep whatever formatting already stood under them.
function statsFormats(
  formats: string[][],
  block: CompsBlock,
  columnCount: number,
): (string | null)[] {
  const last = formats[formats.length - 1] ?? [];
  const numeric = new Set(block.numericColumns);
  return Array.from({ length: columnCount }, (_unused, column) =>
    numeric.has(column) ? (last[column] ?? GENERAL) : null,
  );
}

// One write per statistics column rather than one grid over the whole block:
// a grid would have to carry something for every position, and "General" over
// a column that holds no statistic is a formatting change nobody asked for.
function writeStatsFormats(
  target: Excel.Range,
  formats: (string | null)[],
): void {
  formats.forEach((format, column) => {
    if (format === null) return;
    target.getColumn(column).numberFormat = Array.from(
      { length: STATS_ROWS },
      () => [format],
    );
  });
}

// The two batches the block lands in: the formulas first, so a locked cell is
// named before anything is formatted, then the look each column wears.
async function writeStatsBlock(
  context: Excel.RequestContext,
  target: Excel.Range,
  block: CompsBlock,
  source: { formats: string[][]; columnCount: number },
): Promise<void> {
  target.formulas = statsGrid(block, source.columnCount);
  // A locked sheet refuses the formulas with a host string that names neither
  // the sheet nor the way out; the flow says both itself.
  await syncWrite(context, STAGE);

  writeStatsFormats(
    target,
    statsFormats(source.formats, block, source.columnCount),
  );
  // A null look leaves the cell alone: the columns that hold no numbers keep
  // whatever formatting was already under them.
  writeRuns<PresetLook>(
    target,
    statsPresets(block, source.columnCount),
    (cells, look) => {
      applyPresetFormat(cells.format, look);
    },
  );
  target.select();
  await syncWrite(context, STAGE);
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
    // Asked before the capture, not after: a refusal that had spent an Undo
    // slot would push the modeller's last real action off the five-deep stack.
    if (await sheetProtected(context, sheet)) {
      throw new Error(protectedNote(STAGE));
    }
    await captureUndo(context, target);

    await writeStatsBlock(context, target, block, {
      formats: range.numberFormat as string[][],
      columnCount: range.columnCount,
    });
    return written(block);
  });
}
