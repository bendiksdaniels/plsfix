// Internal API of the src/excel/ folder: private range, fill, workbook-scan,
// defined-name, chart-shell and host-capability helpers no pane code calls
// directly. Exported so sibling section files (and src/excel/links.ts later)
// can import them - the barrel never re-exports this module, so nothing here is
// part of the pane's public surface.

import { type LabelPosition } from "../chart-labels";
import { ANCHOR_PREFIX } from "../link/model";
import { type CellValue } from "../model";
import {
  activeTheme,
  currencyNumberFormat,
  getActiveSettings,
} from "../settings";
import { formatAmount } from "../numbers";
import { brokenNames } from "../workbook";
import { type NumberFormatName } from "./shared";

const staticNumberFormats = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  percent: "0.0%;[Red](0.0%);-",
} as const;

export function numberFormat(name: NumberFormatName): string {
  if (name === "currency") {
    const { currency, language } = getActiveSettings();
    return currencyNumberFormat(currency, language);
  }
  return staticNumberFormats[name];
}

export const SELECTION_CELL_CAP = 5_000;
export const EDIT_CELL_CAP = 500;
// The grid itself: what a block written beside a selection may not run past.
export const SHEET_ROWS = 1_048_576;
export const SHEET_COLUMNS = 16_384;
// What one workbook-wide scan may read in total. The per-sheet cap alone does
// not bound a request: thirty sheets just under it queue a hundred and fifty
// thousand cells into a single sync, which a real model reaches easily and the
// host answers with a bare RequestPayloadSizeLimitExceeded.
export const SCAN_CELL_CAP = SELECTION_CELL_CAP * 4;
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
    return context.workbook.getSelectedRange();
  }

  // Without RangeAreas the count cannot be asked for at all, so the refusal has
  // to be caught where office.js reports it: on the sync after the call.
  const range = context.workbook.getSelectedRange();
  range.load("address");
  try {
    await context.sync();
  } catch (error) {
    const { code } = error as { code?: string };
    if (code !== Excel.ErrorCodes.invalidSelection) throw error;
    throw new Error(`${stage}: select a single range`);
  }
  return range;
}

// A block written beside the selection has to be free first: pls,fix Undo is a
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
export function writeRuns<Key extends string>(
  range: Excel.Range,
  keys: (Key | null)[][],
  write: (block: Excel.Range, key: Key) => void,
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

export interface ScannedSheet {
  index: number;
  name: string;
  range: Excel.Range;
}

// Which sheets a workbook-wide scan can read: an empty sheet has nothing in it,
// a sheet whose used range runs past the per-sheet cap would overflow the
// request payload on its own, and the sheets after the running total passes the
// scan cap would overflow it between them. All three are named as skipped
// rather than quietly left out - the callers render that list, and the style
// scrubber refuses to delete while it is not empty. The ranges come in with
// isNullObject, cellCount, rowIndex and columnIndex already synced; the caller
// loads the grids it needs - values, formulas or both - before the next sync.
export function pickScannableSheets(
  items: Excel.Worksheet[],
  ranges: Excel.Range[],
  cap: number,
  totalCap = SCAN_CELL_CAP,
): { scanned: ScannedSheet[]; skippedSheets: string[] } {
  const scanned: ScannedSheet[] = [];
  const skippedSheets: string[] = [];
  let total = 0;

  ranges.forEach((range, index) => {
    const name = items[index]?.name ?? "";
    if (range.isNullObject) return;
    if (range.cellCount > cap || total + range.cellCount > totalCap) {
      skippedSheets.push(name);
      return;
    }
    total += range.cellCount;
    scanned.push({ index, name, range });
  });

  return { scanned, skippedSheets };
}

export function loadNames(
  context: Excel.RequestContext,
): Excel.NamedItemCollection {
  const names = context.workbook.names;
  names.load("items/name,items/formula");
  return names;
}

// A link anchor whose rows were deleted is a #REF! hidden name by design: the
// Links tab reports it as "Source missing" and owns its removal, and treating
// it as scrub-able here would cut a link the modeller could still heal by
// undoing the delete.
export function brokenIn(names: Excel.NamedItemCollection): string[] {
  return brokenNames(
    names.items
      .filter((item) => !item.name.startsWith(ANCHOR_PREFIX))
      .map((item) => ({
        name: item.name,
        formula: typeof item.formula === "string" ? item.formula : "",
      })),
  );
}

const CHART_TEXT_SIZE = 9;
const CHART_TITLE_SIZE = 12;

// The brand shell every chart gets: our font everywhere, a bold primary title,
// no gridlines, no chart-area frame, legend under the plot.
// The chart surface: the font every label inherits and the corner style.
// Excel for the web does not implement either on its chartex charts (the
// waterfall), and a batch carrying them is rejected whole with
// UnsupportedOperation, so insertWaterfall applies the surface in a batch of
// its own through syncTolerating and keeps the chart when that batch fails.
export function styleChartSurface(chart: Excel.Chart): void {
  const settings = getActiveSettings();
  chart.format.font.name = settings.font;
  chart.format.font.size = CHART_TEXT_SIZE;
  chart.format.font.color = activeTheme().formulaFont;
  chart.format.roundedCorners = false;
}

export function styleChartShell(
  chart: Excel.Chart,
  title: string | null,
  withAxes: boolean,
  surface = true,
): void {
  const settings = getActiveSettings();
  const theme = activeTheme();

  if (surface) styleChartSurface(chart);
  chart.format.border.lineStyle = Excel.ChartLineStyle.none;

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

// The label rule every chart follows: the value and nothing else - no category
// or series name, no percentage, no legend key, no bubble size - in the house
// font, placed where chart-labels.ts says; a null position leaves the host's.
export function styleChartLabels(
  labels: Excel.ChartDataLabels,
  position: LabelPosition | null,
): void {
  const settings = getActiveSettings();
  labels.showValue = true;
  labels.showCategoryName = false;
  labels.showSeriesName = false;
  labels.showPercentage = false;
  labels.showLegendKey = false;
  labels.showBubbleSize = false;
  if (position !== null) labels.position = position;
  labels.format.font.name = settings.font;
  labels.format.font.size = CHART_TEXT_SIZE;
  labels.format.font.color = activeTheme().formulaFont;
}

// Runs the queued batch; a rejection carrying the given error code is
// swallowed and reported as false, anything else is rethrown.
export async function syncTolerating(
  context: Excel.RequestContext,
  code: string,
): Promise<boolean> {
  try {
    await context.sync();
    return true;
  } catch (error) {
    if ((error as { code?: string }).code === code) return false;
    throw error;
  }
}

// Amounts in toasts and labels follow the house style of the pane language.
export function formatChartAmount(value: number): string {
  return formatAmount(value, getActiveSettings().language);
}
