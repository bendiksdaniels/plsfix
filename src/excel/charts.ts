// Charts: the brand shell every chart gets (font, title, legend, no gridlines),
// the native waterfall/bridge builder, chart-selection formatting and the
// floating CAGR label shape.

import { hostSupports } from "./internal";
import { bridgeSeries, cagr, formatCagrLabel } from "../chartmath";
import { type CellValue } from "../model";
import { activeTheme, getActiveSettings, tint } from "../settings";

const BRIDGE_ROW_CAP = 100;
const CHART_TEXT_SIZE = 9;
const CHART_TITLE_SIZE = 12;
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

// The brand shell every chart gets: our font everywhere, a bold primary title,
// no gridlines, no chart-area frame, legend under the plot.
function styleChartShell(
  chart: Excel.Chart,
  title: string | null,
  withAxes: boolean,
): void {
  const settings = getActiveSettings();
  const theme = activeTheme();

  chart.format.font.name = settings.font;
  chart.format.font.size = CHART_TEXT_SIZE;
  chart.format.font.color = theme.formulaFont;
  chart.format.border.lineStyle = Excel.ChartLineStyle.none;
  chart.format.roundedCorners = false;

  if (title !== null) chart.title.text = title;
  chart.title.format.font.name = settings.font;
  chart.title.format.font.size = CHART_TITLE_SIZE;
  chart.title.format.font.bold = true;
  chart.title.format.font.color = settings.primary;

  if (withAxes) {
    for (const axis of [chart.axes.categoryAxis, chart.axes.valueAxis]) {
      axis.format.font.name = settings.font;
      axis.format.font.size = CHART_TEXT_SIZE;
      axis.format.font.color = theme.formulaFont;
      axis.majorGridlines.visible = false;
    }
  }

  chart.legend.position = Excel.ChartLegendPosition.bottom;
  chart.legend.overlay = false;
  chart.legend.format.font.name = settings.font;
  chart.legend.format.font.size = CHART_TEXT_SIZE;
  chart.legend.format.font.color = theme.formulaFont;
}

function formatAmount(value: number): string {
  return value.toLocaleString(undefined, { maximumFractionDigits: 1 });
}

// Native Excel waterfall from a two-column bridge table: labels left, values
// right, first and last rows the opening and closing totals.
export async function insertWaterfall(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values");
    await context.sync();

    if (range.columnCount !== 2 || range.rowCount < 3) {
      throw new Error(
        "Select two columns, labels and values, with three or more rows.",
      );
    }
    if (range.rowCount > BRIDGE_ROW_CAP) {
      throw new Error("A bridge chart supports up to 100 rows.");
    }

    const values: number[] = [];
    for (const row of range.values as CellValue[][]) {
      const value = row[1];
      if (typeof value !== "number" || !Number.isFinite(value)) {
        throw new Error("The second column must hold numbers only.");
      }
      values.push(value);
    }
    const bridge = bridgeSeries(values);

    // The cell above the table is the chart title when it holds text.
    let heading = "Bridge";
    if (range.rowIndex > 0) {
      const above = sheet.getRangeByIndexes(
        range.rowIndex - 1,
        range.columnIndex,
        1,
        1,
      );
      above.load("values");
      await context.sync();
      const text = (above.values as CellValue[][])[0]?.[0];
      if (typeof text === "string" && text.trim()) heading = text.trim();
    }

    const chart = sheet.charts.add(
      Excel.ChartType.waterfall,
      range,
      Excel.ChartSeriesBy.auto,
    );
    styleChartShell(chart, heading, true);
    chart.legend.visible = false;
    chart.dataLabels.showValue = true;

    const series = chart.series.getItemAt(0);
    series.showConnectorLines = true;
    await context.sync();

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

    if (ties) {
      return `Waterfall added: ${values.length} points, ties at ${formatAmount(stated)}`;
    }
    return `Waterfall added: deltas imply ${formatAmount(implied)}, closing total says ${formatAmount(stated)}`;
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
    const periods = (rowCount === 1 ? columnCount : rowCount) - 1;
    if ((rowCount !== 1 && columnCount !== 1) || periods < 1) {
      throw new Error("Select one row or column with at least two numbers.");
    }

    const cells = (range.values as CellValue[][]).flat();
    const first = cells[0];
    const last = cells[cells.length - 1];
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
    return `${label} over ${periods} ${periods === 1 ? "period" : "periods"}`;
  });
}
