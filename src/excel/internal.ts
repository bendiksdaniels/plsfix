// Internal API of the src/excel/ folder: private range, fill, chart-shell and
// host-capability helpers no pane code calls directly. Exported so sibling
// section files (and src/excel/links.ts later) can import them - the barrel
// never re-exports this module, so nothing here is part of the pane's public
// surface.

import { type CellValue } from "../model";
import {
  activeTheme,
  currencyNumberFormat,
  getActiveSettings,
} from "../settings";
import { type NumberFormatName } from "./shared";

const staticNumberFormats = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  percent: "0.0%;[Red](0.0%);-",
} as const;

export function numberFormat(name: NumberFormatName): string {
  if (name === "currency") {
    return currencyNumberFormat(getActiveSettings().currency);
  }
  return staticNumberFormats[name];
}

export const SELECTION_CELL_CAP = 5_000;
export const EDIT_CELL_CAP = 500;
const NO_FILL = "none";
export const BASE_WHITE = "#FFFFFF";

// A whole-column click selects a million cells; reading or writing their grids
// would freeze the pane or overflow the request payload.
export async function withinCap(
  context: Excel.RequestContext,
  range: Excel.Range,
  what: string,
): Promise<Excel.Range> {
  range.load("cellCount");
  await context.sync();
  if (range.cellCount > SELECTION_CELL_CAP) {
    throw new Error(
      `${what} supports up to ${SELECTION_CELL_CAP.toLocaleString()} selected cells at once.`,
    );
  }
  return range;
}

export async function selectionWithinCap(
  context: Excel.RequestContext,
  what: string,
): Promise<Excel.Range> {
  return withinCap(context, context.workbook.getSelectedRange(), what);
}

// getSelectedRange is documented to throw on a multi-area selection (ctrl-click
// two blocks), and it throws as a bare host string with no stage in it. The
// area count is read first so the flow that asked says which one it was.
// getSelectedRanges arrived in ExcelApi 1.9; an older host cannot be asked, and
// falls through to the single-area call it has always made.
export async function selectedSingleRange(
  context: Excel.RequestContext,
  stage: string,
): Promise<Excel.Range> {
  if (hostSupports("1.9")) {
    const areas = context.workbook.getSelectedRanges();
    areas.load("areaCount");
    await context.sync();
    if (areas.areaCount > 1) {
      throw new Error(`${stage}: select a single range`);
    }
  }
  return context.workbook.getSelectedRange();
}

// A block written beside the selection has to be free first: SMT Undo is a
// single slot the modeller has to know to reach for, so a base-case column or a
// comment standing there is not something to overwrite and report afterwards.
export async function requireEmptyBlock(
  context: Excel.RequestContext,
  block: Excel.Range,
  message: string,
): Promise<void> {
  block.load("values");
  await context.sync();
  const occupied = (block.values as CellValue[][]).some((row) =>
    row.some((cell) => cell !== null && cell !== ""),
  );
  if (occupied) throw new Error(message);
}

// One write per run of same-key cells instead of one per cell: model rows are
// usually uniform, so this keeps the batch small on wide selections. A null key
// leaves the cell untouched.
export function writeRuns(
  range: Excel.Range,
  keys: (string | null)[][],
  write: (block: Excel.Range, key: string) => void,
): void {
  keys.forEach((row, rowIndex) => {
    let start = 0;
    while (start < row.length) {
      const key = row[start] ?? null;
      let end = start + 1;
      while (end < row.length && (row[end] ?? null) === key) end += 1;
      if (key !== null) {
        write(
          range.getCell(rowIndex, start).getResizedRange(0, end - start - 1),
          key,
        );
      }
      start = end;
    }
  });
}

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was.
export function fillKey(fill: Excel.CellPropertiesFill | undefined): string {
  const pattern = fill?.pattern ?? Excel.FillPattern.none;
  if (pattern === Excel.FillPattern.none) return NO_FILL;
  return [
    pattern,
    fill?.color ?? BASE_WHITE,
    fill?.patternColor ?? BASE_WHITE,
  ].join("|");
}

export function applyFillKey(block: Excel.Range, key: string): void {
  const { fill } = block.format;
  if (key === NO_FILL) {
    fill.clear();
    return;
  }
  const [pattern, color, patternColor] = key.split("|");
  // Colour first: setting it on an unfilled cell would otherwise force Solid.
  fill.color = color ?? BASE_WHITE;
  fill.pattern = pattern as Excel.FillPattern;
  fill.patternColor = patternColor ?? BASE_WHITE;
}

export function hostSupports(apiSet: string): boolean {
  const requirements = Office.context?.requirements;
  return requirements ? requirements.isSetSupported("ExcelApi", apiSet) : true;
}

const CHART_TEXT_SIZE = 9;
const CHART_TITLE_SIZE = 12;

// The brand shell every chart gets: our font everywhere, a bold primary title,
// no gridlines, no chart-area frame, legend under the plot.
export function styleChartShell(
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

export function formatChartAmount(value: number): string {
  return value.toLocaleString(undefined, { maximumFractionDigits: 1 });
}
