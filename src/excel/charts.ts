// Charts: the native waterfall/bridge builder, chart-selection formatting and
// the floating CAGR label shape. The brand shell they share lives in
// internal.ts, next to the other helpers src/excel/tornado.ts also needs.

import { placeChartBeside, UNPLACED_NOTE } from "./chart-place";
import {
  formatChartAmount,
  hostSupports,
  styleChartShell,
  styleChartSurface,
  syncTolerating,
} from "./internal";
import { bridgeSeries, cagr, formatCagrLabel, seriesSpan } from "../chartmath";
import { type CellValue } from "../model";
import { getActiveSettings, tint } from "../settings";

const BRIDGE_POINT_CAP = 100;
const BRIDGE_MIN_POINTS = 3;
const BRIDGE_TITLE = "Bridge";
const BRIDGE_SHAPE_ERROR =
  "Select labels and values in two adjacent columns, or in two adjacent rows, with three or more points.";
const CAGR_LABEL_WIDTH = 104;
const CAGR_LABEL_HEIGHT = 20;
const CAGR_LABEL_GAP = 8;

// Pies and their relatives have no category or value axis; reading one throws.
const AXIS_FREE_CHARTS = [
  "Pie",
  "PieExploded",
  "PieOfPie",
  "BarOfPie",
  "3DPie",
  "3DPieExploded",
  "Doughnut",
  "DoughnutExploded",
  "Treemap",
  "Sunburst",
  "RegionMap",
];

function chartSeriesColors(): string[] {
  const { primary, accent } = getActiveSettings();
  return [
    accent,
    primary,
    tint(primary, 0.55),
    tint(accent, 0.45),
    tint(primary, 0.78),
    tint(accent, 0.7),
  ];
}

// Two adjacent columns (labels left, values right) or two adjacent rows (labels
// on top, values under them): a bridge reads the same either way, and a
// modeller lays one out both ways.
function bridgeShape(
  rowCount: number,
  columnCount: number,
): { acrossRow: boolean; points: number } {
  if (columnCount === 2 && rowCount >= BRIDGE_MIN_POINTS) {
    return { acrossRow: false, points: rowCount };
  }
  if (rowCount === 2 && columnCount >= BRIDGE_MIN_POINTS) {
    return { acrossRow: true, points: columnCount };
  }
  throw new Error(BRIDGE_SHAPE_ERROR);
}

function bridgeValues(grid: CellValue[][], acrossRow: boolean): number[] {
  const line = acrossRow ? (grid[1] ?? []) : grid.map((row) => row[1] ?? null);
  return line.map((value) => {
    if (typeof value !== "number" || !Number.isFinite(value)) {
      throw new Error(
        acrossRow
          ? "The second row must hold numbers only."
          : "The second column must hold numbers only.",
      );
    }
    return value;
  });
}

// The cell just before the table titles the chart when it holds text: above a
// pair of columns, left of a pair of rows.
async function bridgeHeading(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
  acrossRow: boolean,
): Promise<string> {
  const row = acrossRow ? range.rowIndex : range.rowIndex - 1;
  const column = acrossRow ? range.columnIndex - 1 : range.columnIndex;
  if (row < 0 || column < 0) return BRIDGE_TITLE;

  const cell = sheet.getRangeByIndexes(row, column, 1, 1);
  cell.load("values");
  await context.sync();
  const text = (cell.values as CellValue[][])[0]?.[0];
  return typeof text === "string" && text.trim() ? text.trim() : BRIDGE_TITLE;
}

// Native Excel waterfall from a bridge table: the first and last points are the
// opening and closing totals, everything between them a delta.
export async function insertWaterfall(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values");
    await context.sync();

    const { acrossRow, points: count } = bridgeShape(
      range.rowCount,
      range.columnCount,
    );
    if (count > BRIDGE_POINT_CAP) {
      throw new Error(
        `A bridge chart supports up to ${BRIDGE_POINT_CAP} points.`,
      );
    }

    const values = bridgeValues(range.values as CellValue[][], acrossRow);
    const bridge = bridgeSeries(values);
    const heading = await bridgeHeading(context, sheet, range, acrossRow);

    const chart = sheet.charts.add(
      Excel.ChartType.waterfall,
      range,
      acrossRow ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.auto,
    );
    styleChartShell(chart, heading, true, false);
    chart.legend.visible = false;
    chart.dataLabels.showValue = true;

    const series = chart.series.getItemAt(0);
    series.showConnectorLines = true;
    await context.sync();
    // The geometry travels in its own batch: Excel for the web rejects the one
    // carrying the surface whole, and the placement must not go down with it.
    const placed = await placeChartBeside(context, sheet, chart, range);
    await context.sync();

    // Excel for the web refuses the surface on chartex charts. It is cosmetic,
    // so the waterfall keeps the host's default font there instead of failing.
    styleChartSurface(chart);
    await syncTolerating(context, Excel.ErrorCodes.unsupportedOperation);

    // Office.js has no "set as total" flag for waterfall points, so the opening
    // and closing columns are branded by position instead of by that flag.
    const { primary, accent, external } = getActiveSettings();
    const points = series.points;
    values.forEach((_value, index) => {
      const total = index === 0 || index === values.length - 1;
      const fall = (bridge.fall[index] ?? 0) > 0;
      const color = total ? primary : fall ? external : accent;
      points.getItemAt(index).format.fill.setSolidColor(color);
    });
    await context.sync();

    // Reconciliation: the level the deltas reach against the closing total.
    const last = values.length - 1;
    const implied = (bridge.base[last - 1] ?? 0) + (bridge.rise[last - 1] ?? 0);
    const stated = values[last] ?? 0;
    const ties = Math.abs(implied - stated) <= Math.abs(stated) * 1e-12 + 1e-9;

    const note = placed ? "" : UNPLACED_NOTE;
    if (ties) {
      return `Waterfall added: ${values.length} points, ties at ${formatChartAmount(stated)}${note}`;
    }
    return `Waterfall added: deltas imply ${formatChartAmount(implied)}, closing total says ${formatChartAmount(stated)}${note}`;
  });
}

export async function formatSelectedChart(): Promise<void> {
  await Excel.run(async (context) => {
    const { workbook } = context;
    // The hosted office.js always defines the method, so the host API set decides.
    const callable = (workbook as unknown as Record<string, unknown>)
      .getActiveChartOrNullObject;
    if (typeof callable !== "function" || !hostSupports("1.9")) {
      throw new Error("Chart formatting needs a newer Excel build.");
    }

    const chart = workbook.getActiveChartOrNullObject();
    chart.load("isNullObject");
    await context.sync();
    if (chart.isNullObject) throw new Error("Select a chart first.");

    chart.load("chartType");
    chart.series.load("count");
    await context.sync();

    styleChartShell(
      chart,
      null,
      !AXIS_FREE_CHARTS.includes(String(chart.chartType)),
    );

    const colors = chartSeriesColors();
    for (let index = 0; index < chart.series.count; index += 1) {
      const color = colors[index % colors.length]!;
      chart.series.getItemAt(index).format.fill.setSolidColor(color);
    }
    // One series is already named by the title; its legend is only noise.
    chart.legend.visible = chart.series.count > 1;

    await context.sync();
  });
}

// A floating growth callout beside the series, the way a banker annotates a
// chart by hand. Shapes are worksheet objects, so no range state is touched.
export async function addCagrLabel(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    // Range geometry arrived in 1.10; without it the label lands where Excel
    // drops it and the modeller moves it.
    const positioned = hostSupports("1.10");
    range.load("rowCount,columnCount,values");
    if (positioned) range.load("left,top,width");
    await context.sync();

    const { rowCount, columnCount } = range;
    const cells = (range.values as CellValue[][]).flat();
    // Blank cells at the ends of the line state no period; the series is
    // whatever sits between the first and the last cell that holds something.
    const span = seriesSpan(cells);
    const periods = span ? span.last - span.first : 0;
    if ((rowCount !== 1 && columnCount !== 1) || !span || periods < 1) {
      throw new Error("Select one row or column with at least two numbers.");
    }

    const first = cells[span.first];
    const last = cells[span.last];
    if (typeof first !== "number" || typeof last !== "number") {
      throw new Error("The first and last cells must hold numbers.");
    }

    const shapes = (sheet as unknown as { shapes?: Excel.ShapeCollection })
      .shapes;
    if (
      !shapes ||
      typeof shapes.addTextBox !== "function" ||
      !hostSupports("1.9")
    ) {
      throw new Error("Chart labels need a newer Excel build.");
    }

    const settings = getActiveSettings();
    const label = formatCagrLabel(cagr(first, last, periods));
    const shape = shapes.addTextBox(label);
    shape.width = CAGR_LABEL_WIDTH;
    shape.height = CAGR_LABEL_HEIGHT;
    if (positioned) {
      shape.left = range.left + range.width + CAGR_LABEL_GAP;
      shape.top = range.top;
    }
    shape.fill.clear();
    shape.lineFormat.visible = false;

    const { textFrame } = shape;
    textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.left;
    textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
    const { font } = textFrame.textRange;
    font.name = settings.font;
    font.size = 11;
    font.bold = true;
    font.color = settings.accent;

    await context.sync();
    const note = positioned ? "" : UNPLACED_NOTE;
    return `${label} over ${periods} ${periods === 1 ? "period" : "periods"}${note}`;
  });
}
