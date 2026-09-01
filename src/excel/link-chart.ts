// Reading a chart the way a slide has to redraw it: which kind it is, its
// title, every series' values, the label text the model itself shows and the
// brand colour of each point. Owns the Excel-to-ChartData mapping, and decides
// every string, so the slide side never formats a number. Invariant: a chart it
// cannot read whole is no chart at all, and the picture beside it is the deck's.

import { pieColors, seriesPalette, waterfallColors } from "../chart-colors";
import {
  CHART_MAX_SERIES,
  isChartData,
  type ChartData,
  type ChartKind,
  type ChartSeries,
} from "../link/chart-model";
import {
  activeTheme,
  getActiveSettings,
  type BrandSettings,
} from "../settings";
import { formatChartAmount, hostSupports } from "./internal";
import { parseAddress } from "./shared";

// getDimensionValues: below it a chart link is the picture alone.
const DIMENSION_API = "1.12";
// getDimensionDataSourceString: which cells the values were read from.
const SOURCE_API = "1.15";
// ChartSeries.overlap: what tells a tornado from a plain clustered bar.
const OVERLAP_API = "1.8";
// Excel reports every series of a tornado as fully overlapped.
const FULL_OVERLAP = 100;
// One area of A1, dollars or not: what worksheet.getRange takes.
const A1_RANGE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/;

// First match wins, so the bar families are named before the clustered
// catch-all that would otherwise swallow "BarClustered".
const KIND_RULES: readonly (readonly [RegExp, ChartKind])[] = [
  [/^Waterfall$/, "waterfall"],
  [/^(3D)?Pie(Exploded)?$/, "pie"],
  [/^Line(Markers)?$/, "line"],
  [/^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Stacked/, "stackedBar"],
  [/^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Clustered$/, "bar"],
  [/^(3D)?(Column|Cylinder|Cone|Pyramid)(Col)?.*Stacked/, "stackedColumn"],
  [/Clustered$/, "column"],
];

// What one series answered with, before any of it is a label or a colour.
interface Dimensions {
  name: string;
  categories: string[];
  values: number[];
  source: string;
}

// The chart type and title, read in the batch that carries the picture.
interface Head {
  chartType: string;
  kind: ChartKind;
  title: string | null;
  names: string[];
}

// The kind a slide can draw as shapes, or null for every chart type that stays
// the picture it is today. Columns are laid out side by side, so a clustered
// column whose series fully overlap would draw one bar behind another: that
// chart is one the picture says better.
export function chartKind(
  chartType: string,
  overlap: number | null,
): ChartKind | null {
  const kind =
    KIND_RULES.find(([pattern]) => pattern.test(chartType))?.[1] ?? null;
  return kind === "column" && overlap === FULL_OVERLAP ? null : kind;
}

// A dimension comes back as whatever the host felt like - numbers on the web,
// strings on other builds - so every cell goes through Number(). A blank is a
// gap in the data rather than a zero, and makes the whole chart a picture.
export function valuesOf(raw: (string | number)[]): number[] | null {
  const values: number[] = [];
  for (const cell of raw) {
    const blank = typeof cell === "string" && cell.trim() === "";
    const value = blank ? Number.NaN : Number(cell);
    if (!Number.isFinite(value)) return null;
    values.push(value);
  }
  return values;
}

// What a chart link ships beside its picture, or null when the slide has no
// way to draw it. Always commits at least one sync: the caller queues the
// picture into the same batch, so a chart nothing can be read from costs
// exactly the round trip a chart link costs today.
export async function readChartData(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<ChartData | null> {
  if (!hostSupports(DIMENSION_API)) {
    await context.sync();
    return null;
  }
  const head = await readHead(context, chart);
  if (head === null) return null;
  try {
    return await readBody(context, chart, head);
  } catch {
    // A chart office.js will not describe - a type whose series refuse the
    // dimension reads, a source range that has gone - is one the deck draws
    // as the picture it already has.
    return null;
  }
}

// Sync one, shared with the picture: the type, the title and the series names.
// The overlap waits for the next batch, so a chart type that refuses to answer
// for it can never cost the picture its own round trip.
async function readHead(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<Head | null> {
  chart.load("chartType");
  chart.title.load("text");
  chart.series.load("items/name");
  await context.sync();

  const names = chart.series.items.map((one) => String(one.name ?? ""));
  if (names.length < 1 || names.length > CHART_MAX_SERIES) return null;
  const chartType = String(chart.chartType);
  const kind = chartKind(chartType, null);
  if (kind === null) return null;
  const text = String(chart.title.text ?? "").trim();
  return { chartType, kind, title: text === "" ? null : text, names };
}

// Syncs two and three: every series' categories and values, the cells they came
// from, and the overlap - asked for only where it means something, because the
// chartex types answer for none of it.
async function readBody(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  head: Head,
): Promise<ChartData | null> {
  const wantsOverlap =
    hostSupports(OVERLAP_API) &&
    (head.kind === "column" || head.kind === "bar");
  const results = head.names.map((_name, index) => queueSeries(chart, index));
  if (wantsOverlap) chart.series.load("items/overlap");
  await context.sync();

  const overlaps = wantsOverlap
    ? chart.series.items.map((one) => Number(one.overlap))
    : [];
  const kind = chartKind(head.chartType, overlaps[0] ?? null);
  if (kind === null) return null;
  const dimensions = results.map((one, index) =>
    readSeries(one, head.names[index] ?? ""),
  );
  if (dimensions.some((one) => one === null)) return null;
  const read = dimensions as Dimensions[];
  const labels = await readLabels(context, read);
  const tornado =
    kind === "bar" &&
    overlaps.length > 0 &&
    overlaps.every((one) => one === FULL_OVERLAP);
  return assemble(kind, head.title, read, labels, tornado);
}

// One series' three reads, queued together; the source string only on a host
// that names it, which is what decides whether labels cost a third sync.
function queueSeries(
  chart: Excel.Chart,
  index: number,
): {
  categories: OfficeExtension.ClientResult<string[]>;
  values: OfficeExtension.ClientResult<string[]>;
  source: OfficeExtension.ClientResult<string> | null;
} {
  const series = chart.series.getItemAt(index);
  return {
    categories: series.getDimensionValues("Categories"),
    values: series.getDimensionValues("Values"),
    source: hostSupports(SOURCE_API)
      ? series.getDimensionDataSourceString("Values")
      : null,
  };
}

function readSeries(
  result: ReturnType<typeof queueSeries>,
  name: string,
): Dimensions | null {
  const values = valuesOf(result.values.value);
  if (values === null) return null;
  return {
    name,
    categories: result.categories.value.map(String),
    values,
    source: result.source ? String(result.source.value ?? "") : "",
  };
}

// Sync three: the cells the values were read from, so a label reads exactly as
// the model shows it. A source that is not one plain range - a literal series,
// a defined name, a sheet that has gone - falls back to the house style, and so
// does any cell the range has nothing to say about.
async function readLabels(
  context: Excel.RequestContext,
  series: Dimensions[],
): Promise<string[][]> {
  const house = (): string[][] =>
    series.map((one) => one.values.map(formatChartAmount));
  try {
    const ranges = series.map((one) => sourceRange(context, one.source));
    if (ranges.every((range) => range === null)) return house();
    await context.sync();
    return series.map((one, index) => labelsOf(one.values, ranges[index]));
  } catch {
    return house();
  }
}

function sourceRange(
  context: Excel.RequestContext,
  source: string,
): Excel.Range | null {
  const { sheet, address } = parseAddress(source);
  if (sheet === "" || !A1_RANGE.test(address)) return null;
  const range = context.workbook.worksheets.getItem(sheet).getRange(address);
  range.load("text");
  return range;
}

function labelsOf(
  values: number[],
  range: Excel.Range | null | undefined,
): string[] {
  const text = range ? range.text.flat() : [];
  return values.map((value, index) => {
    const cell = text[index];
    return cell === undefined || cell === "" ? formatChartAmount(value) : cell;
  });
}

// Categories, colours and the brand of the pane, checked against the caps the
// slide draws within: anything the validator refuses ships as no chart at all.
function assemble(
  kind: ChartKind,
  title: string | null,
  series: Dimensions[],
  labels: string[][],
  tornado: boolean,
): ChartData | null {
  const settings = getActiveSettings();
  const points = series[0]?.values.length ?? 0;
  const first = series[0]?.categories ?? [];
  const data: ChartData = {
    v: 1,
    kind,
    title,
    categories: first.length > 0 ? first : numbered(points),
    series: series.map((one, index) =>
      toSeries(one, labels[index] ?? [], kind, index, settings),
    ),
    font: settings.font,
    ink: activeTheme().formulaFont,
    titleColor: settings.primary,
    ...(tornado ? { overlap: true as const } : {}),
  };
  return isChartData(data) ? data : null;
}

function toSeries(
  one: Dimensions,
  labels: string[],
  kind: ChartKind,
  index: number,
  settings: BrandSettings,
): ChartSeries {
  return {
    name: one.name,
    values: one.values,
    labels,
    colors: colorsOf(kind, one.values, index, settings),
  };
}

// The rules the sheet's own charts already follow: a waterfall by position and
// sign, a pie by slice, everything else one palette colour per series.
function colorsOf(
  kind: ChartKind,
  values: number[],
  index: number,
  settings: BrandSettings,
): string[] {
  if (kind === "waterfall") return waterfallColors(values, settings);
  if (kind === "pie") return pieColors(values.length, settings);
  const palette = seriesPalette(settings);
  const color = palette[index % palette.length]!;
  return values.map(() => color);
}

// A chart with no categories of its own numbers its points, the way Excel's own
// axis does.
function numbered(points: number): string[] {
  return Array.from({ length: points }, (_unused, index) => String(index + 1));
}
