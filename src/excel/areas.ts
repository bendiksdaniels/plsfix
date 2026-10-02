// Every rectangle the selection holds. A ctrl-clicked selection is several
// areas, and office.js refuses to hand one over through getSelectedRange at
// all, so the flows that can act area by area ask here instead.
//
// Owns: reopening each area as a range of its own (the addresses come from
// Excel.RangeAreas, ExcelApi 1.9) and the cell cap over all of them together.
// Invariant: the first area is always the one getSelectedRange would have
// served, so a single-area selection behaves exactly as it always did.

import {
  hostSupports,
  overCap,
  SELECTION_CELL_CAP,
  selectionWithinCap,
} from "./internal";
import { parseAddress, splitAreas } from "./shared";

// The areas of a selection sit on one sheet, but the address names it anyway;
// an area with no sheet in it belongs to the one the pane is looking at.
function openArea(context: Excel.RequestContext, area: string): Excel.Range {
  const { sheet, address } = parseAddress(area);
  const sheets = context.workbook.worksheets;
  const worksheet = sheet ? sheets.getItem(sheet) : sheets.getActiveWorksheet();
  return worksheet.getRange(address);
}

// A host below ExcelApi 1.9 has no RangeAreas: it cannot be asked how many
// areas the selection holds, and getSelectedRange answers a ctrl-clicked one
// with a bare host string. The stage and the reason go back in.
async function onlyArea(
  context: Excel.RequestContext,
  what: string,
  capped: boolean,
): Promise<Excel.Range[]> {
  try {
    if (capped) return [await selectionWithinCap(context, what)];
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    return [range];
  } catch (error) {
    const { code } = error as { code?: string };
    if (code === Excel.ErrorCodes.invalidSelection) {
      throw new Error(
        `${what}: this Excel build can only act on one selected block.`,
      );
    }
    throw error;
  }
}

/**
 * The selection as one range per area. A host below ExcelApi 1.9 has no
 * RangeAreas, so it serves the single rectangle it always has.
 */
export async function selectedAreas(
  context: Excel.RequestContext,
  what: string,
): Promise<Excel.Range[]> {
  if (!hostSupports("1.9")) return onlyArea(context, what, false);

  const selection = context.workbook.getSelectedRanges();
  selection.load("areaCount,address");
  await context.sync();

  return splitAreas(selection.address).map((area) => openArea(context, area));
}

/**
 * The same areas, refused as a whole when they hold more cells than one action
 * may read or write. A whole-column click selects a million cells; reading or
 * writing their grids would freeze the pane or overflow the request payload.
 */
export async function cappedAreas(
  context: Excel.RequestContext,
  what: string,
): Promise<Excel.Range[]> {
  if (!hostSupports("1.9")) return onlyArea(context, what, true);

  const ranges = await selectedAreas(context, what);
  for (const range of ranges) range.load("cellCount");
  await context.sync();

  // One area answering -1 (over 2^31-1 cells) would otherwise pull the total
  // below the cap rather than past it.
  const counts = ranges.map((range) => range.cellCount);
  const cells = counts.some((count) => count < 0)
    ? -1
    : counts.reduce((total, count) => total + count, 0);
  if (overCap(cells)) {
    throw new Error(
      `${what} supports up to ${SELECTION_CELL_CAP.toLocaleString()} selected cells at once.`,
    );
  }
  return ranges;
}

/** The cell every cycle reads its current state from: the selection's corner. */
export function activeArea(areas: Excel.Range[]): Excel.Range {
  const first = areas[0];
  if (!first) throw new Error("The selection is empty.");
  return first;
}
