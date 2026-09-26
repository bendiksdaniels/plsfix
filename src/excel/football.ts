// The football field: one floating bar per valuation method, spanning its low
// and high, the way a banker shows what a business is worth. Office charts plot
// ranges only and Excel cannot mix a bar with a marker series, so the floor and
// the band are written to a helper block beside the selection and drawn as a
// stacked bar whose lower series is invisible.
//
// Owns: the Office.js side only; the low/high/range split is pure, in
// src/chartmath.ts, and the helper block itself is shared with the tornado in
// src/excel/chart-blocks.ts. Invariant: the first selected row sits at the top
// of the chart, which is what reversePlotOrder buys.

import { placeChartBeside, UNPLACED_NOTE } from "./chart-place";
import {
  CHART_BLOCK_COLUMNS,
  readTriples,
  requireRoomBeside,
  serialised,
  type TripleRules,
  valueFormat,
  writeHelperBlock,
} from "./chart-blocks";
import {
  hostSupports,
  selectedSingleRange,
  styleChartShell,
  styleChartSurface,
  syncTolerating,
  withinCap,
} from "./internal";
import { seriesPalette } from "../chart-colors";
import { footballField } from "../chartmath";
import { type CellValue } from "../model";
import { getActiveSettings } from "../settings";

const STAGE = "Football field";
const FOOTBALL_ROW_CAP = 20;
const FOOTBALL_MIN_ROWS = 2;
const FOOTBALL_TITLE = "Valuation range";
const FOOTBALL_HEADERS = ["Method", "Low", "Range"];
const SHAPE_ERROR = `${STAGE}: select at least three columns - label, low and high.`;
// Only the first three columns are plotted: Excel cannot mix a bar with a
// marker series, so a fourth column - a point estimate, usually - is ignored.
const EXTRA_COLUMNS_NOTE = "; only label, low and high are used";
// reversePlotOrder is ExcelApi 1.7 and the axis number format 1.8. Without
// them the chart is still a football field, drawn in Excel's own row order
// with the host's own axis labels, which is worth saying out loud.
const BASIC_AXES_NOTE = "; plain axes on this build";
// Excel for the web refuses chart.format.font and roundedCorners on this
// chart the same way it refuses them on a chartex chart (lessons 29.08); the
// tolerated batch below keeps the chart and its placement when that happens,
// and this is the only visible trace left for the modeller.
const SURFACE_NOTE = "; some styling could not be applied";

const FOOTBALL_RULES: TripleRules = {
  minRows: FOOTBALL_MIN_ROWS,
  tooFew: `${STAGE}: need at least two rows`,
  notNumbers: `${STAGE}: the low and high columns must hold numbers`,
  rowCap: {
    max: FOOTBALL_ROW_CAP,
    message: `${STAGE} supports up to ${FOOTBALL_ROW_CAP} rows.`,
  },
};

function styleFootball(chart: Excel.Chart, format: string): void {
  // The surface stays out of this batch: Excel for the web refuses the font
  // and the corners on this chart, and that refusal must not take the rest
  // of the styling, the chart itself or its placement down with it. Applied
  // afterwards, in its own tolerated batch (runFootballField).
  styleChartShell(chart, FOOTBALL_TITLE, true, false);
  chart.legend.visible = false;

  // Series 1 is the invisible floor up to each method's low; series 2 is the
  // band a reader actually sees.
  const floor = chart.series.getItemAt(0);
  floor.format.fill.clear();
  floor.format.line.lineStyle = Excel.ChartLineStyle.none;
  chart.series
    .getItemAt(1)
    .format.fill.setSolidColor(seriesPalette(getActiveSettings())[0]!);

  // A bar chart plots the first category at the bottom; reversing the order
  // puts the first selected method on top, the way a football field is read.
  if (hostSupports("1.7")) chart.axes.categoryAxis.reversePlotOrder = true;
  if (hostSupports("1.8")) chart.axes.valueAxis.numberFormat = format;
}

function swapNote(swaps: number): string {
  if (swaps === 0) return "";
  const rows = swaps === 1 ? "row" : "rows";
  return `; ${String(swaps)} ${rows} had low above high, swapped`;
}

function notes(
  swaps: number,
  columnCount: number,
  placed: boolean,
  surfaced: boolean,
): string {
  return [
    swapNote(swaps),
    columnCount > CHART_BLOCK_COLUMNS ? EXTRA_COLUMNS_NOTE : "",
    hostSupports("1.7") && hostSupports("1.8") ? "" : BASIC_AXES_NOTE,
    placed ? "" : UNPLACED_NOTE,
    surfaced ? "" : SURFACE_NOTE,
  ].join("");
}

/**
 * Label, low and high in three columns. The helper block lands immediately
 * right of the selection, and pls,fix Undo captures whatever stood there first.
 */
async function runFootballField(
  context: Excel.RequestContext,
): Promise<string> {
  // The cap answers before the values are asked for: a clicked column header
  // is a million cells, and the row cap only runs after the read.
  const range = await withinCap(
    context,
    await selectedSingleRange(context, STAGE),
    STAGE,
  );
  const sheet = range.worksheet;
  range.load("columnCount,rowIndex,columnIndex,values,numberFormat");
  await context.sync();

  if (range.columnCount < CHART_BLOCK_COLUMNS) throw new Error(SHAPE_ERROR);
  requireRoomBeside(range, STAGE);

  const rows = readTriples(range.values as CellValue[][], FOOTBALL_RULES);
  const field = footballField(rows);
  const format = valueFormat(range.numberFormat as string[][], rows.length);
  const block = await writeHelperBlock(context, sheet, range, {
    stage: STAGE,
    headers: FOOTBALL_HEADERS,
    rows: field.labels.map((label, index) => [
      label,
      field.low[index] ?? 0,
      field.range[index] ?? 0,
    ]),
    format,
  });

  const chart = sheet.charts.add(
    Excel.ChartType.barStacked,
    block,
    Excel.ChartSeriesBy.columns,
  );
  styleFootball(chart, format);
  await context.sync();
  const placed = await placeChartBeside(context, sheet, chart, block);
  await context.sync();

  // Its own batch, tolerated the way the waterfall's own surface is: a
  // refusal here must never undo the placement that already landed.
  styleChartSurface(chart);
  const surfaced = await syncTolerating(
    context,
    Excel.ErrorCodes.unsupportedOperation,
  );

  const tail = notes(field.swaps, range.columnCount, placed, surfaced);
  return `${STAGE} added: ${String(rows.length)} ranges${tail}`;
}

// Serialised with the tornado: a second press must meet this press's helper
// block, not an empty one, so it hits the "not empty" refusal instead of
// drawing a second chart on top of the first.
export async function insertFootballField(): Promise<string> {
  return serialised(() => Excel.run(runFootballField));
}
