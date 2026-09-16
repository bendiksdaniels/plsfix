// The horizontal bar and tornado-row primitives: one row per category, series
// stacked or overlapped within it. Split out of chart-shapes.ts to keep both
// it and chart-shapes-parts.ts under the line cap. Pure: no Office.js.
// Invariant: every primitive lies inside the box it was given.

import {
  BAR_FILL,
  valueX,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";
import { barRange, boxAt, rect, segmentLabel } from "./chart-shapes-parts";

// The row order for one category: input order, or (a tornado) the widest
// swing first regardless of sign, so the shorter bar behind it stays visible.
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

// Bar and tornado rows: one row per category, series stacked vertically
// within it (or overlapping, longer first, when `overlap` is set).
export function barBars(
  data: ChartData,
  chart: Box,
  plot: Box,
  scale: ValueScale,
): Primitive[] {
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
      // A stacked row is one band, like a stacked column's one slot.
      const top =
        overlap || stacked
          ? plot.top + i * rowSlot + (rowSlot - barHeight) / 2
          : plot.top +
            i * rowSlot +
            (rowSlot - data.series.length * barHeight) / 2 +
            slotIndex * barHeight;
      const left = valueX(lo, scale, plot);
      const box = boxAt(left, top, valueX(hi, scale, plot) - left, barHeight);
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      out.push(...segmentLabel(data, series, j, i, box, stacked, false, chart));
    });
  });
  return out;
}
