// The sensitivity tornado: drivers ranked by how far they move the answer, each
// drawn as a bar spanning its low and high around the base. Office charts plot
// ranges only and the deltas are not in the model, so they are written to a
// helper block beside the selection and charted from there.

import { placeChartBeside, UNPLACED_NOTE } from "./chart-place";
import {
  formatChartAmount,
  hostSupports,
  requireEmptyBlock,
  selectedSingleRange,
  SHEET_COLUMNS,
  styleChartLabels,
  styleChartShell,
} from "./internal";
import { captureUndo } from "./undo";
import { type TornadoDriver, tornadoSeries } from "../chartmath";
import { type CellValue } from "../model";
import { getActiveSettings } from "../settings";

const TORNADO_COLUMNS = 3;
const TORNADO_ROW_CAP = 100;
// Both halves of a driver share one bar row, with the rows drawn close together.
const TORNADO_OVERLAP = 100;
const TORNADO_GAP_WIDTH = 40;
const TORNADO_TITLE = "Sensitivity";
const TORNADO_HEADERS = ["Driver", "Low", "High"];
const TORNADO_SHAPE_ERROR =
  "tornado: need 3 columns (label, low, high) and at least 2 rows";
// Bar overlap and gap width arrived in ExcelApi 1.8; without them the tornado
// is drawn as a plain clustered bar chart, which is worth saying out loud.
const BASIC_BARS_NOTE = "; plain bars on this build";

function isFiniteNumber(value: CellValue | undefined): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

// The header row is optional and detected, not declared: a first row whose two
// outcome cells hold no numbers is a pair of column titles, never a driver.
function readDrivers(grid: CellValue[][]): TornadoDriver[] {
  const first = grid[0] ?? [];
  const headed = !isFiniteNumber(first[1]) && !isFiniteNumber(first[2]);
  const body = headed ? grid.slice(1) : grid;
  if (body.length < 2) throw new Error(TORNADO_SHAPE_ERROR);

  return body.map((row) => {
    const [label, low, high] = row;
    if (!isFiniteNumber(low) || !isFiniteNumber(high)) {
      throw new Error("tornado: the low and high columns must hold numbers");
    }
    return { label: label === null ? "" : String(label), low, high };
  });
}

// The helper block's deltas share the outcomes' unit, so they wear the format
// of the first driver's low cell; a headed selection has one row above it.
function outcomeFormat(formats: string[][], driverCount: number): string {
  const firstDriverRow = formats.length - driverCount;
  return formats[firstDriverRow]?.[1] ?? "General";
}

// The block's header row and label column stay plain; the two delta columns
// carry the outcomes' format, which the chart's labels read off them.
function blockFormats(format: string, driverCount: number): string[][] {
  return [
    TORNADO_HEADERS.map(() => "General"),
    ...Array.from({ length: driverCount }, () => ["General", format, format]),
  ];
}

interface TornadoHeader {
  heading: string;
  base: number | null;
}

// One read of the row above the selection answers both questions it can: text
// over the label column titles the chart, and the first number over the outcome
// columns is the base case. Without either, the defaults stand.
async function readTornadoHeader(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
): Promise<TornadoHeader> {
  if (range.rowIndex === 0) return { heading: TORNADO_TITLE, base: null };

  const above = sheet.getRangeByIndexes(
    range.rowIndex - 1,
    range.columnIndex,
    1,
    TORNADO_COLUMNS,
  );
  above.load("values");
  await context.sync();

  const cells = (above.values as CellValue[][])[0] ?? [];
  const title = cells[0];
  return {
    heading:
      typeof title === "string" && title.trim() ? title.trim() : TORNADO_TITLE,
    base: cells.slice(1).find(isFiniteNumber) ?? null,
  };
}

function styleTornado(chart: Excel.Chart, heading: string): void {
  styleChartShell(chart, heading, true);
  styleChartLabels(chart.dataLabels, "OutsideEnd");
  // A bar chart plots the first category at the bottom; reversing the order
  // puts the widest swing on top, which is the shape a tornado is read by.
  if (hostSupports("1.7")) chart.axes.categoryAxis.reversePlotOrder = true;

  const { accent, external } = getActiveSettings();
  // Downside in the same colour a bridge paints a fall, upside in the accent.
  [external, accent].forEach((color, index) => {
    const series = chart.series.getItemAt(index);
    series.format.fill.setSolidColor(color);
    if (hostSupports("1.8")) {
      series.overlap = TORNADO_OVERLAP;
      series.gapWidth = TORNADO_GAP_WIDTH;
    }
  });
}

// Label, low outcome, high outcome; the helper block lands immediately right of
// the selection, and pls,fix Undo captures whatever stood there first.
export async function insertTornado(): Promise<string> {
  return Excel.run(async (context) => {
    const range = await selectedSingleRange(context, "tornado");
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values,numberFormat");
    await context.sync();

    if (range.columnCount !== TORNADO_COLUMNS || range.rowCount < 2) {
      throw new Error(TORNADO_SHAPE_ERROR);
    }
    if (range.rowCount > TORNADO_ROW_CAP) {
      throw new Error(`tornado: supports up to ${TORNADO_ROW_CAP} drivers`);
    }
    if (
      range.columnIndex + range.columnCount + TORNADO_COLUMNS >
      SHEET_COLUMNS
    ) {
      throw new Error("tornado: no room to the right of the selection");
    }

    const drivers = readDrivers(range.values as CellValue[][]);
    const format = outcomeFormat(
      range.numberFormat as string[][],
      drivers.length,
    );
    const { heading, base } = await readTornadoHeader(context, sheet, range);
    const series = tornadoSeries(drivers, base);

    const block = sheet.getRangeByIndexes(
      range.rowIndex,
      range.columnIndex + range.columnCount,
      series.labels.length + 1,
      TORNADO_COLUMNS,
    );
    await requireEmptyBlock(
      context,
      block,
      "tornado: cells to the right of the selection are not empty",
    );
    await captureUndo(context, block);
    block.values = [
      TORNADO_HEADERS,
      ...series.labels.map((label, index) => [
        label,
        series.low[index] ?? 0,
        series.high[index] ?? 0,
      ]),
    ];
    block.numberFormat = blockFormats(format, series.labels.length);
    await context.sync();

    const chart = sheet.charts.add(
      Excel.ChartType.barClustered,
      block,
      Excel.ChartSeriesBy.columns,
    );
    styleTornado(chart, heading);
    await context.sync();
    const placed = await placeChartBeside(context, sheet, chart, block);
    await context.sync();

    const count = series.labels.length;
    const notes = [
      placed ? "" : UNPLACED_NOTE,
      hostSupports("1.7") && hostSupports("1.8") ? "" : BASIC_BARS_NOTE,
    ].join("");
    return `Tornado added: ${count} drivers, base ${formatChartAmount(series.base)}${notes}`;
  });
}
