// "Select consistent region": the block of cells the active cell's formula was
// filled into, handed back as the selection. Owns the one read of the sheet's
// used range and the growth call into the pure audit module. Invariant: the
// flow only reads and selects, so it never goes through syncWrite.

import { scanCapSentence } from "./audit";
import { overCap, selectedSingleRange } from "./internal";
import { consistentRegion } from "../audit";
import { type CellValue } from "../model";

const STAGE = "Consistent region";
const NO_FORMULA = "The active cell has no formula.";
const ONLY_ONE = "Only the active cell has this formula.";
// The overlay's cap over the overlay's own wording, this tool's name in front.
const OVER_CAP = scanCapSentence("Select consistent region");

/**
 * Grows the rectangle of cells that share the active cell's R1C1 formula and
 * selects it. Read-only: four syncs (the area count, the used range's size,
 * its grid, the select), and no pls,fix Undo slot, because nothing is written.
 */
export async function selectConsistentRegion(): Promise<string> {
  return Excel.run(async (context) => {
    // Which cell is active is only meaningful for one block; a ctrl-clicked
    // selection is refused before anything is read.
    await selectedSingleRange(context, STAGE);

    const cell = context.workbook.getActiveCell();
    cell.load("rowIndex,columnIndex");
    const sheet = cell.worksheet;
    const used = sheet.getUsedRangeOrNullObject(true);
    // The cap is counted before the grid is asked for: a sheet used down to
    // the last row would otherwise queue a million cells into one sync.
    used.load("isNullObject,cellCount,rowIndex,columnIndex");
    await context.sync();

    if (used.isNullObject) return NO_FORMULA;
    if (overCap(used.cellCount)) throw new Error(OVER_CAP);

    used.load("formulasR1C1");
    await context.sync();

    const region = consistentRegion(
      used.formulasR1C1 as CellValue[][],
      cell.rowIndex - used.rowIndex,
      cell.columnIndex - used.columnIndex,
    );
    if (region === null) return NO_FORMULA;

    sheet
      .getRangeByIndexes(
        used.rowIndex + region.row,
        used.columnIndex + region.column,
        region.rowCount,
        region.columnCount,
      )
      .select();
    await context.sync();

    const cells = region.rowCount * region.columnCount;
    if (cells === 1) return ONLY_ONE;
    return `${cells.toLocaleString()} cells share this formula`;
  });
}
