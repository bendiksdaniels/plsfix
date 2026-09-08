// The primitive builders every chart kind shares (box/rect/line/text
// factories, truncation, the stacked-segment cursor, row ordering) plus pie
// wedges and the legend; split out of chart-shapes.ts to keep both files
// under the line cap. Pure: no Office.js, nothing from src/excel or src/ppt.
// Invariant: every primitive lies inside the box it was given.

import type { BridgeSeries } from "./chartmath";
import type { Box } from "./layout";
import {
  normalizeAngle,
  type ChartData,
  type ChartSeries,
} from "./link/chart-model";
import {
  LABEL_HEIGHT,
  LABEL_PAD,
  LABEL_SIZE,
  LEGEND_BAND,
  MIN_SEGMENT,
  PIE_LABEL_RADIUS,
  PIE_RADIUS,
  SWATCH,
  TITLE_SIZE,
  darkFill,
  textWidth,
  type Ellipse,
  type Line,
  type Primitive,
  type Rect,
  type Text,
  type Wedge,
} from "./chart-shapes";

const ELLIPSIS = "…";
const PIE_START_DEGREES = -90; // 12 o'clock, in PowerPoint's clockwise-from-3-o'clock convention
const DEGREES_PER_TURN = 360;
const DEGREES_TO_RADIANS = Math.PI / 180;

// Compact builders: every call site names a box, a colour and a name instead
// of repeating the full object shape (this is what keeps both chart-shapes
// files under the line cap despite prettier's 80-column wrap).
export function boxAt(
  left: number,
  top: number,
  width: number,
  height: number,
): Box {
  return { left, top, width, height };
}
export function rect(box: Box, color: string, name: string): Rect {
  return { kind: "rect", box, color, name };
}
export function lineShape(
  box: Box,
  color: string,
  weight: number,
  name: string,
): Line {
  return { kind: "line", box, color, weight, rising: false, name };
}

// A connector between two points, normalised so width and height are never
// negative (PowerPoint throws InvalidArgument on a negative side, on both the
// add and a later write) and flagged when the segment climbs left to right,
// since that box alone no longer says which way the line ran.
export function lineBetween(
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  color: string,
  weight: number,
  name: string,
): Line {
  return {
    kind: "line",
    box: boxAt(
      Math.min(x0, x1),
      Math.min(y0, y1),
      Math.abs(x1 - x0),
      Math.abs(y1 - y0),
    ),
    color,
    weight,
    rising: y1 < y0,
    name,
  };
}
export function label(
  box: Box,
  text: string,
  color: string,
  align: Text["align"],
  name: string,
): Text {
  return {
    kind: "text",
    box,
    text,
    size: LABEL_SIZE,
    bold: false,
    color,
    align,
    name,
  };
}
export function titleText(box: Box, text: string, color: string): Text {
  return {
    kind: "text",
    box,
    text,
    size: TITLE_SIZE,
    bold: true,
    color,
    align: "l",
    name: "title",
  };
}
function wedgeShape(
  box: Box,
  start: number,
  end: number,
  color: string,
  name: string,
): Wedge {
  return { kind: "wedge", box, start, end, color, name };
}
function ellipseShape(box: Box, color: string, name: string): Ellipse {
  return { kind: "ellipse", box, color, name };
}

export function truncate(text: string, size: number, maxWidth: number): string {
  if (textWidth(text, size) <= maxWidth) return text;
  let cut = text;
  while (cut.length > 0 && textWidth(cut + ELLIPSIS, size) > maxWidth)
    cut = cut.slice(0, -1);
  return cut.length > 0 ? cut + ELLIPSIS : ELLIPSIS;
}

// A stacked series' running position: the value range [lo, hi] a category's
// next positive or negative segment occupies, advancing the shared cursor.
export function stackSegment(
  cursor: { pos: number; neg: number },
  value: number,
): [number, number] {
  if (value >= 0) {
    const lo = cursor.pos;
    cursor.pos += value;
    return [lo, cursor.pos];
  }
  const hi = cursor.neg;
  cursor.neg += value;
  return [cursor.neg, hi];
}

export function barRange(
  stacked: boolean,
  cursor: { pos: number; neg: number },
  value: number,
): [number, number] {
  if (!stacked) return [Math.min(0, value), Math.max(0, value)];
  return stackSegment(cursor, value);
}

// A waterfall bar's value range, and which side of it the label belongs on.
// The range is ordered low to high because a bar that ends below where it
// started - a closing total under zero - otherwise comes out with a negative
// height, which PowerPoint refuses outright.
export function bridgeRange(
  bridge: BridgeSeries,
  i: number,
): { lo: number; hi: number; below: boolean } {
  const from = bridge.base[i]!;
  const to = from + bridge.rise[i]! + bridge.fall[i]!;
  return {
    lo: Math.min(from, to),
    hi: Math.max(from, to),
    below: bridge.fall[i]! > 0 || to < from,
  };
}

export function rowOrder(
  series: ChartSeries[],
  i: number,
  overlap: boolean,
): number[] {
  const indices = series.map((_, j) => j);
  if (!overlap) return indices;
  return indices.sort(
    (a, b) => Math.abs(series[b]!.values[i]!) - Math.abs(series[a]!.values[i]!),
  );
}

// A value label outside the bar it belongs to: above a positive column bar or
// right of a positive bar-chart bar, the far side for a negative one. A
// stacked segment's label sits centred inside instead, white on a dark fill,
// and is dropped when the segment is too thin to hold it.
export function segmentLabel(
  data: ChartData,
  series: ChartSeries,
  j: number,
  i: number,
  value: number,
  box: Box,
  stacked: boolean,
  vertical: boolean,
): Text[] {
  const content = series.labels[i]!;
  const name = `label ${j}.${i}`;
  if (stacked) {
    if ((vertical ? box.height : box.width) < MIN_SEGMENT) return [];
    const color = darkFill(series.colors[i]!) ? "#FFFFFF" : data.ink;
    return [label(box, content, color, "c", name)];
  }
  if (vertical) {
    const top = value >= 0 ? box.top - LABEL_HEIGHT : box.top + box.height;
    return [
      label(
        boxAt(box.left, top, box.width, LABEL_HEIGHT),
        content,
        data.ink,
        "c",
        name,
      ),
    ];
  }
  const width = textWidth(content, LABEL_SIZE);
  const left = value >= 0 ? box.left + box.width : box.left - width;
  const top = box.top + (box.height - LABEL_HEIGHT) / 2;
  return [
    label(
      boxAt(left, top, width, LABEL_HEIGHT),
      content,
      data.ink,
      value >= 0 ? "l" : "r",
      name,
    ),
  ];
}

export function legendItems(data: ChartData): string[] {
  return data.kind === "pie" ? data.categories : data.series.map((s) => s.name);
}

// Greedily wraps legend entries into rows that fit `width`, one row minimum.
export function packLegendRows(items: string[], width: number): number[][] {
  const rows: number[][] = [];
  let row: number[] = [];
  let used = 0;
  items.forEach((item, i) => {
    const w = SWATCH + LABEL_PAD + textWidth(item, LABEL_SIZE) + LABEL_PAD;
    if (row.length > 0 && used + w > width) {
      rows.push(row);
      row = [];
      used = 0;
    }
    row.push(i);
    used += w;
  });
  if (row.length > 0) rows.push(row);
  return rows.length > 0 ? rows : [[]];
}

function colorOf(data: ChartData, i: number): string {
  return data.kind === "pie"
    ? data.series[0]!.colors[i]!
    : data.series[i]!.colors[0]!;
}

// Legend rows wrapped under the plot: a square swatch beside the series name,
// or the category name for a pie, packed by the same rule `bands` reserved
// space for. Swatches are rects named "legend swatch": a bar is a rect named
// "bar", which is how a reader tells the two apart.
export function legend(data: ChartData, band: Box): Primitive[] {
  const items = legendItems(data);
  // Only the rows the band was actually given: in a box too small to hold
  // every entry the legend is cut short rather than drawn past the chart.
  const rows = packLegendRows(items, band.width).slice(
    0,
    Math.floor(band.height / LEGEND_BAND),
  );
  const out: Primitive[] = [];
  rows.forEach((row, rowIndex) => {
    let x = band.left;
    const top = band.top + rowIndex * LEGEND_BAND;
    row.forEach((i) => {
      const swatch = boxAt(x, top + (LEGEND_BAND - SWATCH) / 2, SWATCH, SWATCH);
      out.push(rect(swatch, colorOf(data, i), `legend swatch ${i}`));
      x += SWATCH + LABEL_PAD;
      const width = textWidth(items[i]!, LABEL_SIZE);
      out.push(
        label(
          boxAt(x, top, width, LEGEND_BAND),
          items[i]!,
          data.ink,
          "l",
          `legend ${i}`,
        ),
      );
      x += width + LABEL_PAD;
    });
  });
  return out;
}

// A slice's label, anchored left of centre when it points right and right of
// centre when it points left, so the text grows away from the pie.
function sliceLabel(
  content: string,
  ink: string,
  cx: number,
  cy: number,
  r: number,
  midDegrees: number,
  name: string,
): Text {
  const rad = midDegrees * DEGREES_TO_RADIANS;
  const px = cx + r * PIE_LABEL_RADIUS * Math.cos(rad);
  const py = cy + r * PIE_LABEL_RADIUS * Math.sin(rad);
  const width = textWidth(content, LABEL_SIZE);
  const align = Math.cos(rad) >= 0 ? "l" : "r";
  const left = align === "l" ? px : px - width;
  return label(
    boxAt(left, py - LABEL_HEIGHT / 2, width, LABEL_HEIGHT),
    content,
    ink,
    align,
    name,
  );
}

// Slices from 12 o'clock clockwise, sized by share of the positive total;
// zero and negative values are skipped and do not consume any angle. A lone
// 100% slice is an ellipse, since PowerPoint's adjustments cannot sweep 360.
export function pieWedges(data: ChartData, plot: Box): Primitive[] {
  const series = data.series[0]!;
  const cx = plot.left + plot.width / 2;
  const cy = plot.top + plot.height / 2;
  const r = (Math.min(plot.width, plot.height) / 2) * PIE_RADIUS;
  const box = boxAt(cx - r, cy - r, r * 2, r * 2);
  const total = series.values.reduce((sum, v) => (v > 0 ? sum + v : sum), 0);
  if (total <= 0) return [];

  const drawn = data.categories
    .map((_, i) => i)
    .filter((i) => series.values[i]! > 0);
  if (drawn.length === 1) {
    return [ellipseShape(box, series.colors[drawn[0]!]!, `slice ${drawn[0]}`)];
  }

  const out: Primitive[] = [];
  let cumulative = 0;
  drawn.forEach((i) => {
    const span = (series.values[i]! / total) * DEGREES_PER_TURN;
    const start = normalizeAngle(PIE_START_DEGREES + cumulative);
    const end = normalizeAngle(PIE_START_DEGREES + cumulative + span);
    out.push(wedgeShape(box, start, end, series.colors[i]!, `slice ${i}`));
    const mid = PIE_START_DEGREES + cumulative + span / 2;
    out.push(
      sliceLabel(series.labels[i]!, data.ink, cx, cy, r, mid, `label ${i}`),
    );
    cumulative += span;
  });
  return out;
}
