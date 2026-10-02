// The chart data a picture link may carry: the kinds a slide can draw, the
// caps past which the picture is the honest answer and the words that say so,
// the validator every decoder runs, and the angle rule PowerPoint's pie
// adjustments follow.
// Pure: no Office.js, nothing from src/excel or src/ppt.

export type ChartKind =
  | "column"
  | "stackedColumn"
  | "bar"
  | "stackedBar"
  | "waterfall"
  | "pie"
  | "line";

// One series: a value, its displayed text and its brand colour per category.
export interface ChartSeries {
  name: string;
  values: number[];
  labels: string[];
  colors: string[];
}

export interface ChartData {
  v: 1;
  kind: ChartKind;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
  font: string;
  ink: string;
  titleColor: string;
  // A tornado: both series share one row, drawn longest first.
  overlap?: true;
}

export const CHART_MAX_POINTS = 40;
export const CHART_MAX_SERIES = 6;
export const CHART_MAX_SLICES = 12;
export const CHART_MIN_POINTS = 2;
export const CHART_TITLE_MAX = 80;

const KINDS: readonly string[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "waterfall",
  "pie",
  "line",
];
const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const FULL_TURN = 360;
const HALF_TURN = 180;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStrings(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((v) => typeof v === "string");
}

function isSeries(value: unknown, points: number): value is ChartSeries {
  if (!isRecord(value) || typeof value.name !== "string") return false;
  const { values, labels, colors } = value;
  return (
    Array.isArray(values) &&
    values.length === points &&
    values.every((v) => typeof v === "number" && Number.isFinite(v)) &&
    isStrings(labels) &&
    labels.length === points &&
    isStrings(colors) &&
    colors.length === points &&
    colors.every((color) => HEX_COLOR.test(color))
  );
}

function withinTitle(title: unknown): boolean {
  return (
    title === null ||
    (typeof title === "string" && title.length <= CHART_TITLE_MAX)
  );
}

// Shape, consistency and the caps in one pass: a payload that fails here
// simply has no chart data, and the picture beside it is what gets drawn.
export function isChartData(value: unknown): value is ChartData {
  if (!isRecord(value) || value.v !== 1) return false;
  if (typeof value.kind !== "string" || !KINDS.includes(value.kind)) {
    return false;
  }
  if (!withinTitle(value.title) || !isStrings(value.categories)) return false;
  const points = value.categories.length;
  if (points < CHART_MIN_POINTS || points > CHART_MAX_POINTS) return false;
  const { series } = value;
  if (!Array.isArray(series) || series.length < 1) return false;
  if (series.length > CHART_MAX_SERIES) return false;
  if (!series.every((one) => isSeries(one, points))) return false;
  if (value.kind === "pie") {
    if (series.length !== 1 || points > CHART_MAX_SLICES) return false;
  }
  if (typeof value.font !== "string" || typeof value.ink !== "string") {
    return false;
  }
  if (typeof value.titleColor !== "string") return false;
  return value.overlap === undefined || value.overlap === true;
}

// PowerPoint stores a pie's start and end as degrees clockwise from 3 o'clock
// and reads any angle back on (-180, 180]: 200 becomes -160, 360 becomes 0.
export function normalizeAngle(degrees: number): number {
  const wrapped =
    ((((degrees + HALF_TURN) % FULL_TURN) + FULL_TURN) % FULL_TURN) - HALF_TURN;
  return wrapped === -HALF_TURN ? HALF_TURN : wrapped;
}

// The reason a picture link has no chart data, in the words the panes show
// after "as a picture": the first cap the chart is outside of, counted the way
// the modeller sees the chart.
export function seriesCountIssue(count: number): string | null {
  if (count < 1) return "no series";
  if (count > CHART_MAX_SERIES) {
    return `${String(count)} series; shapes draw up to ${String(CHART_MAX_SERIES)}`;
  }
  return null;
}

export function chartCapIssue(
  kind: ChartKind,
  points: number,
  seriesCount: number,
): string | null {
  const series = seriesCountIssue(seriesCount);
  if (series !== null) return series;
  if (points < CHART_MIN_POINTS) {
    const noun = points === 1 ? "point" : "points";
    return `${String(points)} ${noun}; shapes need at least ${String(CHART_MIN_POINTS)}`;
  }
  if (points > CHART_MAX_POINTS) {
    return `${String(points)} points; shapes draw up to ${String(CHART_MAX_POINTS)}`;
  }
  if (kind === "pie" && seriesCount !== 1) {
    return `a pie with ${String(seriesCount)} series; pie shapes draw one`;
  }
  if (kind === "pie" && points > CHART_MAX_SLICES) {
    return `${String(points)} slices; pie shapes draw up to ${String(CHART_MAX_SLICES)}`;
  }
  return null;
}

// The one host reason that is not a capability or a cap: PowerPoint took the
// draw batch and never answered it, so the shapes it had already taken were
// deleted again and the picture went where the chart would have been.
export const CHART_HOST_SILENT =
  "PowerPoint stopped answering while drawing the shapes";

// The sentence both panes show for a chart that stays a picture: the insert
// note in PowerPoint and the export line in Excel repeat it word for word.
export function pictureNote(issue: string): string {
  return `as a picture: ${issue}`;
}
