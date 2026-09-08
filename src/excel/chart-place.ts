// Where a chart the add-in inserts lands. Excel drops a new chart wherever the
// view happens to be, which on a model sheet means on top of the numbers, so
// every chart this add-in adds is moved onto a block of cells big enough for it
// that holds no values and covers no other chart.
//
// Owns: the candidate order (right of the anchor, below it, below everything
// the sheet uses) and the two syncs it takes to answer. Invariant: a chart is
// only moved onto a block whose own used range came back empty.

import { hostSupports, SHEET_COLUMNS, SHEET_ROWS } from "./internal";
import { type Box, dropBelow, overlaps } from "../layout";

// Points kept clear between two charts, and the cells of air between a chart
// and the block it was placed beside.
const CHART_GAP = 12;
const GAP_ROWS = 1;
const GAP_COLUMNS = 1;

// What a chart toast adds when the host cannot say where a range sits, so the
// modeller knows the chart landed wherever Excel dropped it.
export const UNPLACED_NOTE = "; Excel placed it";

// Excel's own defaults, in points. RangeFormat.rowHeight and columnWidth
// answer null when the range's rows or columns are not all one size - a label
// column beside a value column never is - and the office.js types say number,
// so the null arrives untyped and would size the block at one point per cell.
const DEFAULT_ROW_HEIGHT = 15;
const DEFAULT_COLUMN_WIDTH = 48;

function sizeOr(value: number | null, fallback: number): number {
  return typeof value === "number" && value > 0 ? value : fallback;
}

interface Corner {
  row: number;
  column: number;
}

interface Plan {
  corners: Corner[];
  rows: number;
  columns: number;
  others: Box[];
}

interface Slot {
  box: Box;
  free: boolean;
}

function boxOf(range: Excel.Range): Box {
  return {
    left: range.left,
    top: range.top,
    width: range.width,
    height: range.height,
  };
}

// Right of the anchor, below the anchor, below everything the sheet uses - each
// one cell clear of what it sits beside, and dropped when it would run off the
// grid.
function candidateCorners(
  anchor: Excel.Range,
  sheetBottom: number,
  rows: number,
  columns: number,
): Corner[] {
  return [
    {
      row: anchor.rowIndex,
      column: anchor.columnIndex + anchor.columnCount + GAP_COLUMNS,
    },
    {
      row: anchor.rowIndex + anchor.rowCount + GAP_ROWS,
      column: anchor.columnIndex,
    },
    { row: sheetBottom + GAP_ROWS, column: anchor.columnIndex },
  ].filter(
    (corner) =>
      corner.row + rows <= SHEET_ROWS &&
      corner.column + columns <= SHEET_COLUMNS,
  );
}

// One sync: the anchor's size in cells and in points, the chart's own size, the
// charts already on the sheet and how far down the sheet's own data reaches.
async function readPlan(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chart: Excel.Chart,
  anchor: Excel.Range,
): Promise<Plan> {
  anchor.load(
    "rowIndex,columnIndex,rowCount,columnCount,format/rowHeight,format/columnWidth",
  );
  chart.load("name,width,height");
  const charts = sheet.charts;
  charts.load("items/name,items/left,items/top,items/width,items/height");
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("isNullObject,rowIndex,rowCount");
  await context.sync();

  const rowHeight = sizeOr(anchor.format.rowHeight, DEFAULT_ROW_HEIGHT);
  const columnWidth = sizeOr(anchor.format.columnWidth, DEFAULT_COLUMN_WIDTH);
  const rows = Math.max(1, Math.ceil(chart.height / rowHeight));
  const columns = Math.max(1, Math.ceil(chart.width / columnWidth));
  const bottom = used.isNullObject
    ? anchor.rowIndex + anchor.rowCount
    : used.rowIndex + used.rowCount;

  return {
    rows,
    columns,
    corners: candidateCorners(anchor, bottom, rows, columns),
    others: charts.items
      .filter((other) => other.name !== chart.name)
      .map((other) => ({
        left: other.left,
        top: other.top,
        width: other.width,
        height: other.height,
      })),
  };
}

// One sync: each candidate block's own used range - values only, so a block
// that is merely formatted still counts as free - and its box in points.
async function readSlots(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  plan: Plan,
): Promise<Slot[]> {
  const blocks = plan.corners.map((corner) =>
    sheet.getRangeByIndexes(corner.row, corner.column, plan.rows, plan.columns),
  );
  const used = blocks.map((block) => block.getUsedRangeOrNullObject(true));
  for (const block of blocks) block.load("left,top,width,height");
  for (const range of used) range.load("isNullObject");
  await context.sync();

  return blocks.map((block, index) => ({
    box: boxOf(block),
    free: used[index]?.isNullObject ?? false,
  }));
}

function pick(slots: Slot[], others: readonly Box[]): Box | null {
  for (const slot of slots) {
    const clear = others.every(
      (other) => !overlaps(slot.box, other, CHART_GAP),
    );
    if (slot.free && clear) return slot.box;
  }
  // The last candidate sits below everything the sheet uses, so no value can be
  // under it; only another chart can still be in the way, and it steps under it.
  const last = slots[slots.length - 1];
  return last ? dropBelow(last.box, others, CHART_GAP) : null;
}

/**
 * Moves the chart onto the first free block beside its anchor. False when the
 * host cannot say where a range sits (range geometry is ExcelApi 1.10), which
 * leaves the chart where Excel dropped it; the caller says so in its toast.
 */
export async function placeChartBeside(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  chart: Excel.Chart,
  anchor: Excel.Range,
): Promise<boolean> {
  if (!hostSupports("1.10")) return false;

  const plan = await readPlan(context, sheet, chart, anchor);
  const box = pick(await readSlots(context, sheet, plan), plan.others);
  if (!box) return false;

  chart.left = box.left;
  chart.top = box.top;
  return true;
}
