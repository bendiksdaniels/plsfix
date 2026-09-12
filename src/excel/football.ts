// The football field: one floating bar per valuation method, spanning its low
// and high, the way a banker shows what a business is worth. Office charts plot
// ranges only and Excel cannot mix a bar with a marker series, so the floor and
// the band are written to a helper block beside the selection and drawn as a
// stacked bar whose lower series is invisible.
//
// Owns: the Office.js side only; the low/high/range split is pure, in
// src/chartmath.ts. Invariant: the first selected row sits at the top of the
// chart, which is what reversePlotOrder buys.

import { placeChartBeside, UNPLACED_NOTE } from "./chart-place";
import {
  hostSupports,
  requireEmptyBlock,
  selectedSingleRange,
  SHEET_COLUMNS,
  styleChartShell,
  withinCap,
} from "./internal";
import { syncWrite } from "./protection";
import { captureUndo } from "./undo";
import { seriesPalette } from "../chart-colors";
import { footballField, type FootballRow } from "../chartmath";
import { type CellValue } from "../model";
import { getActiveSettings } from "../settings";

const STAGE = "Football field";
const FOOTBALL_COLUMNS = 3;
export const FOOTBALL_ROW_CAP = 20;
const FOOTBALL_MIN_ROWS = 2;
const FOOTBALL_TITLE = "Valuation range";
const FOOTBALL_HEADERS = ["Method", "Low", "Range"];
const SHAPE_ERROR = `${STAGE}: select at least three columns - label, low and high.`;
const NUMBERS_ERROR = `${STAGE}: the low and high columns must hold numbers`;
// Only the first three columns are plotted: Excel cannot mix a bar with a
// marker series, so a fourth column - a point estimate, usually - is ignored.
const EXTRA_COLUMNS_NOTE = "; only label, low and high are used";
// reversePlotOrder is ExcelApi 1.7 and the axis number format 1.8. Without
// them the chart is still a football field, drawn in Excel's own row order
// with the host's own axis labels, which is worth saying out loud.
const BASIC_AXES_NOTE = "; plain axes on this build";

function isFiniteNumber(value: CellValue | undefined): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

// The header row is detected, not declared: a first row whose low and high
// cells hold no numbers is a set of column titles, never a valuation method.
function readRows(grid: CellValue[][]): FootballRow[] {
  const first = grid[0] ?? [];
  const headed = !isFiniteNumber(first[1]) && !isFiniteNumber(first[2]);
  const body = headed ? grid.slice(1) : grid;
  if (body.length < FOOTBALL_MIN_ROWS) {
    throw new Error(`${STAGE}: need at least two rows`);
  }
  if (body.length > FOOTBALL_ROW_CAP) {
    throw new Error(`${STAGE} supports up to ${FOOTBALL_ROW_CAP} rows.`);
  }

  return body.map((row) => {
    const [label, low, high] = row;
    if (!isFiniteNumber(low) || !isFiniteNumber(high)) {
      throw new Error(NUMBERS_ERROR);
    }
    return { label: label === null ? "" : String(label), low, high };
  });
}

// The floor and the band share the valuation's unit, so both wear the format of
// the first row's low cell; a headed selection has one row above it.
function lowFormat(formats: string[][], rowCount: number): string {
  const firstRow = formats.length - rowCount;
  return formats[firstRow]?.[1] ?? "General";
}

// The header row and the label column stay plain; the two number columns carry
// the valuation's format, which the value axis reads off them.
function blockFormats(format: string, rowCount: number): string[][] {
  return [
    FOOTBALL_HEADERS.map(() => "General"),
    ...Array.from({ length: rowCount }, () => ["General", format, format]),
  ];
}

function styleFootball(chart: Excel.Chart, format: string): void {
  styleChartShell(chart, FOOTBALL_TITLE, true);
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

/**
 * Label, low and high in three columns. The helper block lands immediately
 * right of the selection, and pls,fix Undo captures whatever stood there first.
 */
export async function insertFootballField(): Promise<string> {
  return Excel.run(async (context) => {
    // The cap answers before the values are asked for: a clicked column header
    // is a million cells, and the row cap only runs after the read.
    const range = await withinCap(
      context,
      await selectedSingleRange(context, STAGE),
      STAGE,
    );
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values,numberFormat");
    await context.sync();

    if (range.columnCount < FOOTBALL_COLUMNS) throw new Error(SHAPE_ERROR);
    if (
      range.columnIndex + range.columnCount + FOOTBALL_COLUMNS >
      SHEET_COLUMNS
    ) {
      throw new Error(`${STAGE}: no room to the right of the selection`);
    }

    const rows = readRows(range.values as CellValue[][]);
    const field = footballField(rows);
    const format = lowFormat(range.numberFormat as string[][], rows.length);

    const block = sheet.getRangeByIndexes(
      range.rowIndex,
      range.columnIndex + range.columnCount,
      rows.length + 1,
      FOOTBALL_COLUMNS,
    );
    await requireEmptyBlock(
      context,
      block,
      `${STAGE}: cells to the right of the selection are not empty`,
    );
    await captureUndo(context, block);
    block.values = [
      FOOTBALL_HEADERS,
      ...field.labels.map((label, index) => [
        label,
        field.low[index] ?? 0,
        field.range[index] ?? 0,
      ]),
    ];
    block.numberFormat = blockFormats(format, rows.length);
    // A locked sheet refuses the helper block by name, before a chart is added
    // over a block that never landed.
    await syncWrite(context, STAGE);

    const chart = sheet.charts.add(
      Excel.ChartType.barStacked,
      block,
      Excel.ChartSeriesBy.columns,
    );
    styleFootball(chart, format);
    await context.sync();
    const placed = await placeChartBeside(context, sheet, chart, block);
    await context.sync();

    const notes = [
      swapNote(field.swaps),
      range.columnCount > FOOTBALL_COLUMNS ? EXTRA_COLUMNS_NOTE : "",
      hostSupports("1.7") && hostSupports("1.8") ? "" : BASIC_AXES_NOTE,
      placed ? "" : UNPLACED_NOTE,
    ].join("");
    return `${STAGE} added: ${String(rows.length)} ranges${notes}`;
  });
}
