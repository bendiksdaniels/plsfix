// Turns ChartData plus a box into drawable primitives (rect, wedge, ellipse,
// text, line) for column, bar, waterfall and pie charts. Pure: no Office.js,
// nothing from src/excel or src/ppt; every number is in the caller's points.
// Invariant: every primitive lies inside the box it was given.

import { bridgeSeries } from "./chartmath";
import type { Box, Size } from "./layout";
import type { ChartData, ChartKind } from "./link/chart-model";
import {
  barRange,
  boxAt,
  label,
  legend,
  legendItems,
  lineShape,
  packLegendRows,
  pieWedges,
  rect,
  rowOrder,
  segmentLabel,
  titleText,
  truncate,
} from "./chart-shapes-parts";

export interface Rect {
  kind: "rect";
  box: Box;
  color: string;
  name: string;
}
export interface Wedge {
  kind: "wedge";
  box: Box;
  start: number;
  end: number;
  color: string;
  name: string;
}
export interface Ellipse {
  kind: "ellipse";
  box: Box;
  color: string;
  name: string;
}
export interface Text {
  kind: "text";
  box: Box;
  text: string;
  size: number;
  bold: boolean;
  color: string;
  align: "l" | "c" | "r";
  name: string;
}
export interface Line {
  kind: "line";
  box: Box;
  color: string;
  weight: number;
  name: string;
}
export type Primitive = Rect | Wedge | Ellipse | Text | Line;

export const LABEL_SIZE = 9;
export const TITLE_SIZE = 12;
export const TITLE_BAND = 18;
export const LEGEND_BAND = 16;
export const CATEGORY_BAND = 14;
export const BAR_FILL = 0.6;
export const LABEL_HEIGHT = 18;
export const LABEL_PAD = 7;
export const LINE_WEIGHT = 0.75;
export const PIE_RADIUS = 0.8;
export const PIE_LABEL_RADIUS = 1.18;
export const CHAR_WIDTH = 0.55;
export const SWATCH = 8;
export const BAR_LABEL_COLUMN = 0.28;
export const MIN_SEGMENT = 12;
export const MIN_SIZE: Size = { width: 200, height: 120 };

const LUMINANCE_R = 0.299;
const LUMINANCE_G = 0.587;
const LUMINANCE_B = 0.114;
const DARK_THRESHOLD = 0.5;
const FULL_BYTE = 255;

// Text width is estimated, not measured: a fixed points-per-character factor
// plus the padding a text box's default insets add.
export function textWidth(text: string, size: number): number {
  return CHAR_WIDTH * size * text.length + LABEL_PAD;
}

// Perceived brightness (ITU-R BT.601 luma weights) of a "#RRGGBB" colour.
export function darkFill(hex: string): boolean {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const luminance =
    (LUMINANCE_R * r + LUMINANCE_G * g + LUMINANCE_B * b) / FULL_BYTE;
  return luminance < DARK_THRESHOLD;
}

// Scales the source image size down to fit maxWidth keeping its aspect ratio;
// never returns smaller than MIN_SIZE in either dimension.
export function chartSize(source: Size, maxWidth: number): Size {
  const width = Math.max(Math.min(source.width, maxWidth), MIN_SIZE.width);
  const aspect = source.height / source.width;
  const height = Math.max(width * aspect, MIN_SIZE.height);
  return { width, height };
}

interface Bands {
  title: Box;
  plot: Box;
  categoryBand: Box | null;
  legendBand: Box | null;
  labelColumn: Box | null;
}

// Carves the chart box into its fixed bands (title, legend, category axis or
// label column) and returns what is left as the plot.
function bands(data: ChartData, box: Box): Bands {
  const barLike = data.kind === "bar" || data.kind === "stackedBar";
  const isPie = data.kind === "pie";
  const title = boxAt(box.left, box.top, box.width, TITLE_BAND);
  let bottom = box.top + box.height;

  let legendBand: Box | null = null;
  if (data.series.length > 1 || isPie) {
    const rows = packLegendRows(legendItems(data), box.width).length;
    const height = rows * LEGEND_BAND;
    legendBand = boxAt(box.left, bottom - height, box.width, height);
    bottom -= height;
  }

  let categoryBand: Box | null = null;
  let left = box.left;
  let width = box.width;
  if (!barLike && !isPie) {
    categoryBand = boxAt(
      box.left,
      bottom - CATEGORY_BAND,
      box.width,
      CATEGORY_BAND,
    );
    bottom -= CATEGORY_BAND;
  }

  let labelColumn: Box | null = null;
  const plotTop = box.top + TITLE_BAND;
  if (barLike) {
    const colWidth = box.width * BAR_LABEL_COLUMN;
    labelColumn = boxAt(box.left, plotTop, colWidth, bottom - plotTop);
    left = box.left + colWidth;
    width = box.width - colWidth;
  }

  const plot = boxAt(left, plotTop, width, bottom - plotTop);
  return { title, plot, categoryBand, legendBand, labelColumn };
}

interface ValueScale {
  min: number;
  max: number;
}

// The reach of the chart's values: from the smallest negative sum to the
// largest positive sum, always including zero for the baseline.
function valueScale(data: ChartData): ValueScale {
  if (data.kind === "waterfall") {
    const bridge = bridgeSeries(data.series[0]!.values);
    let min = 0,
      max = 0;
    bridge.base.forEach((lo, i) => {
      const hi = lo + bridge.rise[i]! + bridge.fall[i]!;
      min = Math.min(min, lo, hi);
      max = Math.max(max, lo, hi);
    });
    return { min, max };
  }
  const stacked = data.kind === "stackedColumn" || data.kind === "stackedBar";
  let min = 0,
    max = 0;
  data.categories.forEach((_, i) => {
    if (stacked) {
      let pos = 0,
        neg = 0;
      data.series.forEach((s) => {
        if (s.values[i]! >= 0) pos += s.values[i]!;
        else neg += s.values[i]!;
      });
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    } else {
      data.series.forEach((s) => {
        max = Math.max(max, s.values[i]!);
        min = Math.min(min, s.values[i]!);
      });
    }
  });
  return { min, max };
}

function valueY(value: number, scale: ValueScale, plot: Box): number {
  const range = scale.max - scale.min || 1;
  return plot.top + ((scale.max - value) / range) * plot.height;
}

function valueX(value: number, scale: ValueScale, plot: Box): number {
  const range = scale.max - scale.min || 1;
  return plot.left + ((value - scale.min) / range) * plot.width;
}

function baselineLine(
  kind: ChartKind,
  data: ChartData,
  plot: Box,
  scale: ValueScale,
): Line {
  if (kind === "bar" || kind === "stackedBar") {
    const box = boxAt(valueX(0, scale, plot), plot.top, 0, plot.height);
    return lineShape(box, data.ink, LINE_WEIGHT, "baseline");
  }
  const box = boxAt(plot.left, valueY(0, scale, plot), plot.width, 0);
  return lineShape(box, data.ink, LINE_WEIGHT, "baseline");
}

// Clustered or stacked columns: one slot per category, series side by side or
// concatenated top to bottom; series outer so a stacked category's segments
// (same left) sit next to each other in the output.
function columnBars(
  data: ChartData,
  plot: Box,
  scale: ValueScale,
): Primitive[] {
  const stacked = data.kind === "stackedColumn";
  const n = data.categories.length;
  const slot = plot.width / n;
  const cursors = data.categories.map(() => ({ pos: 0, neg: 0 }));
  const out: Primitive[] = [];
  data.series.forEach((series, j) => {
    data.categories.forEach((_, i) => {
      const value = series.values[i]!;
      const barWidth = stacked
        ? BAR_FILL * slot
        : (BAR_FILL * slot) / data.series.length;
      const [lo, hi] = barRange(stacked, cursors[i]!, value);
      const left = stacked
        ? plot.left + i * slot + (slot - barWidth) / 2
        : plot.left +
          i * slot +
          (slot - data.series.length * barWidth) / 2 +
          j * barWidth;
      const top = valueY(hi, scale, plot);
      const box = boxAt(left, top, barWidth, valueY(lo, scale, plot) - top);
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      out.push(...segmentLabel(data, series, j, i, value, box, stacked, true));
    });
  });
  return out;
}

// Bar and tornado rows: one row per category, series stacked vertically
// within it (or overlapping, longer first, when `overlap` is set).
function barBars(data: ChartData, plot: Box, scale: ValueScale): Primitive[] {
  const stacked = data.kind === "stackedBar";
  const overlap = data.overlap === true && !stacked;
  const n = data.categories.length;
  const rowSlot = plot.height / n;
  const out: Primitive[] = [];
  data.categories.forEach((_, i) => {
    const order = rowOrder(data.series, i, overlap);
    const cursor = { pos: 0, neg: 0 };
    const barHeight =
      overlap || stacked
        ? BAR_FILL * rowSlot
        : (BAR_FILL * rowSlot) / data.series.length;
    order.forEach((j, slotIndex) => {
      const series = data.series[j]!;
      const value = series.values[i]!;
      const [lo, hi] = barRange(stacked, cursor, value);
      const top = overlap
        ? plot.top + i * rowSlot + (rowSlot - barHeight) / 2
        : plot.top +
          i * rowSlot +
          (rowSlot - data.series.length * barHeight) / 2 +
          slotIndex * barHeight;
      const left = valueX(lo, scale, plot);
      const box = boxAt(left, top, valueX(hi, scale, plot) - left, barHeight);
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      out.push(...segmentLabel(data, series, j, i, value, box, stacked, false));
    });
  });
  return out;
}

// Floating bars from bridgeSeries (a bar's value range is [base, base+rise+
// fall]), joined by a connector at the running level after each bar.
function waterfallBars(
  data: ChartData,
  plot: Box,
  scale: ValueScale,
): Primitive[] {
  const series = data.series[0]!;
  const bridge = bridgeSeries(series.values);
  const n = data.categories.length;
  const slot = plot.width / n;
  const barWidth = BAR_FILL * slot;
  const out: Primitive[] = [];
  const runningLevel: number[] = [];
  data.categories.forEach((_, i) => {
    const lo = bridge.base[i]!;
    const hi = lo + bridge.rise[i]! + bridge.fall[i]!;
    const left = plot.left + i * slot + (slot - barWidth) / 2;
    const top = valueY(hi, scale, plot);
    const box = boxAt(left, top, barWidth, valueY(lo, scale, plot) - top);
    out.push(rect(box, series.colors[i]!, `bar 0.${i}`));
    const isFall = bridge.fall[i]! > 0;
    const labelTop = isFall ? box.top + box.height : box.top - LABEL_HEIGHT;
    const labelBox = boxAt(left, labelTop, barWidth, LABEL_HEIGHT);
    out.push(label(labelBox, series.labels[i]!, data.ink, "c", `label 0.${i}`));
    runningLevel.push(valueY(bridge.base[i]! + bridge.rise[i]!, scale, plot));
  });
  for (let i = 0; i < n - 1; i += 1) {
    const fromRight = plot.left + i * slot + (slot + barWidth) / 2;
    const toLeft = plot.left + (i + 1) * slot + (slot - barWidth) / 2;
    const box = boxAt(fromRight, runningLevel[i]!, toLeft - fromRight, 0);
    out.push(lineShape(box, data.ink, LINE_WEIGHT, `connector ${i}`));
  }
  return out;
}

// One label per category: centred under its slot (column/waterfall) or
// right-aligned in the left column (bar/tornado); truncated to fit.
function categoryLabels(
  data: ChartData,
  plot: Box,
  band: Box,
  barLike: boolean,
): Text[] {
  const n = data.categories.length;
  return data.categories.map((category, i) => {
    if (barLike) {
      const rowSlot = plot.height / n;
      const content = truncate(category, LABEL_SIZE, band.width);
      const box = boxAt(band.left, plot.top + i * rowSlot, band.width, rowSlot);
      return label(box, content, data.ink, "r", `category ${i}`);
    }
    const slot = plot.width / n;
    const content = truncate(category, LABEL_SIZE, slot);
    const box = boxAt(plot.left + i * slot, band.top, slot, band.height);
    return label(box, content, data.ink, "c", `category ${i}`);
  });
}

export function layoutChart(data: ChartData, box: Box): Primitive[] {
  const { title, plot, categoryBand, legendBand, labelColumn } = bands(
    data,
    box,
  );
  const out: Primitive[] = [];
  if (data.title !== null)
    out.push(titleText(title, data.title, data.titleColor));
  if (data.kind === "pie") {
    out.push(...pieWedges(data, plot));
  } else {
    const scale = valueScale(data);
    if (data.kind === "column" || data.kind === "stackedColumn") {
      out.push(
        ...columnBars(data, plot, scale),
        baselineLine(data.kind, data, plot, scale),
      );
      if (categoryBand)
        out.push(...categoryLabels(data, plot, categoryBand, false));
    } else if (data.kind === "bar" || data.kind === "stackedBar") {
      out.push(
        ...barBars(data, plot, scale),
        baselineLine(data.kind, data, plot, scale),
      );
      if (labelColumn)
        out.push(...categoryLabels(data, plot, labelColumn, true));
    } else {
      out.push(
        ...waterfallBars(data, plot, scale),
        baselineLine(data.kind, data, plot, scale),
      );
      if (categoryBand)
        out.push(...categoryLabels(data, plot, categoryBand, false));
    }
  }
  if (legendBand) out.push(...legend(data, legendBand));
  return out;
}
