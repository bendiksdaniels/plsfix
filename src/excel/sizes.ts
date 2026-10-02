// Row-height and column-width cycles: pressing the same button or shortcut
// again steps the selected band through a fixed ladder of sizes (src/cycles.ts,
// in points). Two deliberate differences from the other format cycles:
//
//   * No pls,fix Undo. Row heights and column widths are sheet state, not cell
//     state, and `getCellProperties` carries neither - so `captureUndo` could
//     not put a size back and is not called. These two actions run outside the
//     undo net by design, which README "Known limits" states.
//   * No cell cap. The work is one scalar read plus one band write however many
//     cells are selected, and clicking the row or column headers - a selection
//     far over the cap - is the normal way to reach these. The multi-area guard
//     still applies: `getSelectedRange` throws on a ctrl-clicked selection.

import { selectedSingleRange } from "./internal";
import { syncWrite } from "./protection";
import { buildSizeCycles, nextSize } from "../cycles";

// The first row of the selection carries the cycle state, the way the active
// cell carries a format cycle's. Read from `getRow(0)` rather than the range:
// Excel reports null for a range whose rows differ in height.
export async function applyRowHeightCycle(): Promise<string> {
  return Excel.run(async (context) => {
    const range = await selectedSingleRange(context, "Row height");
    const first = range.getRow(0);
    first.load("format/rowHeight");
    await context.sync();

    const next = nextSize(first.format.rowHeight, buildSizeCycles().rowHeight);
    // The whole rows the selection touches, not the selected cells: a height
    // belongs to the row, and Excel would widen the band on its own anyway.
    range.getEntireRow().format.rowHeight = next;
    await syncWrite(context, "Row height");
    return `Row height ${next} pt (outside pls,fix Undo)`;
  });
}

export async function applyColumnWidthCycle(): Promise<string> {
  return Excel.run(async (context) => {
    const range = await selectedSingleRange(context, "Column width");
    const first = range.getColumn(0);
    first.load("format/columnWidth");
    await context.sync();

    const next = nextSize(
      first.format.columnWidth,
      buildSizeCycles().columnWidth,
    );
    range.getEntireColumn().format.columnWidth = next;
    await syncWrite(context, "Column width");
    return `Column width ${next} (outside pls,fix Undo)`;
  });
}
