// Direct proof of lineSeries, called on its own (not through layoutChart, so
// the scale is hand-set rather than derived): a marker centred in its slot
// and sized between the min and max off the slot width, rising set only for
// a segment that actually climbs (flat and falling both false), the label
// above a marker or below it once the marker sits on the plot's top edge, and
// every primitive non-negative and inside the box at the 40-point/6-series
// caps. chart-shapes.test.ts already proves this through layoutChart with one
// small dataset; this file drives the module directly with its own fixtures.
import { describe, expect, it } from "vitest";
import { lineSeries } from "./chart-shapes-line";
import {
  TITLE_BAND,
  valueY,
  type Ellipse,
  type Line,
  type Text,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  type ChartData,
  type ChartSeries,
} from "./link/chart-model";

const SCALE: ValueScale = { min: 0, max: 20 };

function lineData(
  categories: string[],
  series: Omit<ChartSeries, "colors">[],
): ChartData {
  return {
    v: 1,
    kind: "line",
    title: null,
    categories,
    series: series.map((s) => ({
      ...s,
      colors: s.values.map(() => "#2EC4B6"),
    })),
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
}

function markers(out: ReturnType<typeof lineSeries>): Ellipse[] {
  return out.filter((p): p is Ellipse => p.kind === "ellipse");
}
function connectors(out: ReturnType<typeof lineSeries>): Line[] {
  return out.filter((p): p is Line => p.kind === "line");
}
function labels(out: ReturnType<typeof lineSeries>): Text[] {
  return out.filter((p): p is Text => p.kind === "text");
}

function insideBox(inner: Box, outer: Box, slack = 0.01): boolean {
  return (
    inner.width >= 0 &&
    inner.height >= 0 &&
    inner.left >= outer.left - slack &&
    inner.top >= outer.top - slack &&
    inner.left + inner.width <= outer.left + outer.width + slack &&
    inner.top + inner.height <= outer.top + outer.height + slack
  );
}

describe("lineSeries marker placement", () => {
  it("centres every marker on its slot, away from the edges a wide plot never clamps", () => {
    const plot: Box = { left: 1_000, top: 500, width: 400, height: 100 };
    const data = lineData(
      ["A", "B", "C", "D"],
      [{ name: "s", values: [10, 5, 15, 10], labels: ["10", "5", "15", "10"] }],
    );
    const out = lineSeries(data, plot, SCALE);
    const points = markers(out);
    expect(points).toHaveLength(4);

    const slot = plot.width / data.categories.length;
    points.forEach((marker, i) => {
      const expectedCentre = plot.left + (i + 0.5) * slot;
      expect(marker.box.left + marker.box.width / 2).toBeCloseTo(
        expectedCentre,
        6,
      );
    });
  });

  it("sizes the marker off the slot width, clamped between the min and max", () => {
    const scale: ValueScale = { min: 0, max: 100 };
    const widthAt = (slotWidth: number): number => {
      // One category, so the slot is the whole plot width.
      const plot: Box = { left: 0, top: 0, width: slotWidth, height: 200 };
      const data = lineData(
        ["A"],
        [{ name: "s", values: [50], labels: ["50"] }],
      );
      return markers(lineSeries(data, plot, scale))[0]!.box.width;
    };

    expect(widthAt(10)).toBe(3); // 10 / 5 = 2, clamped up to MARKER_MIN
    expect(widthAt(25)).toBe(5); // 25 / 5 = 5, inside the range as-is
    expect(widthAt(100)).toBe(6); // 100 / 5 = 20, clamped down to MARKER_MAX
  });
});

describe("lineSeries rising", () => {
  it("flags rising only for the segment that actually climbs, not a flat or a falling one", () => {
    const plot: Box = { left: 0, top: 500, width: 400, height: 100 };
    // 10 -> 10 (flat) -> 20 (climbs, and lands exactly on the plot's own top
    // edge, the scale max) -> 5 (falls back down).
    const data = lineData(
      ["A", "B", "C", "D"],
      [{ name: "s", values: [10, 10, 20, 5], labels: ["10", "10", "20", "5"] }],
    );
    const lines = connectors(lineSeries(data, plot, SCALE));
    expect(lines).toHaveLength(3);
    expect(lines.map((l) => l.rising)).toEqual([false, true, false]);
    expect(lines.every((l) => l.box.width >= 0 && l.box.height >= 0)).toBe(
      true,
    );
  });
});

describe("lineSeries labels", () => {
  it("sits a label above its marker normally, and below it once the marker touches the plot's top edge", () => {
    const plot: Box = { left: 0, top: 500, width: 300, height: 100 };
    // Point 0 sits mid-scale, well clear of the top edge; point 1 is the
    // scale max, whose marker clamps to the plot's own top (valueY(20) = plot.top).
    const data = lineData(
      ["A", "B"],
      [{ name: "s", values: [10, 20], labels: ["10", "20"] }],
    );
    const out = lineSeries(data, plot, SCALE);
    const [midMarker, topMarker] = markers(out);
    const [midLabel, topLabel] = labels(out);

    expect(topMarker!.box.top).toBeCloseTo(plot.top, 6); // confirms the edge case is real
    expect(midLabel!.box.top + midLabel!.box.height).toBeCloseTo(
      midMarker!.box.top,
      6,
    );
    expect(topLabel!.box.top).toBeCloseTo(
      topMarker!.box.top + topMarker!.box.height,
      6,
    );
  });
});

describe("lineSeries at the 40-point, 6-series cap", () => {
  it("keeps every marker, connector and label non-negative and inside the plot", () => {
    const plot: Box = { left: 20, top: 40, width: 900, height: 300 };
    const categories = Array.from(
      { length: CHART_MAX_POINTS },
      (_unused, i) => `C${String(i)}`,
    );
    const scale: ValueScale = { min: -50, max: 50 };
    const series = Array.from({ length: CHART_MAX_SERIES }, (_unused, j) => {
      const values = categories.map((_unused2, i) => {
        const magnitude = ((i + j) % 50) + 1;
        return i % 2 === 0 ? magnitude : -magnitude;
      });
      return { name: `s${String(j)}`, values, labels: values.map(String) };
    });

    const data = lineData(categories, series);
    const out = lineSeries(data, plot, scale);

    expect(markers(out)).toHaveLength(CHART_MAX_POINTS * CHART_MAX_SERIES);
    expect(connectors(out)).toHaveLength(
      (CHART_MAX_POINTS - 1) * CHART_MAX_SERIES,
    );
    expect(labels(out)).toHaveLength(CHART_MAX_POINTS * CHART_MAX_SERIES);
    // A label above a marker that sits close to, but not exactly on, the top
    // edge pokes above `plot` itself by design: layoutChart always gives
    // lineSeries a plot with a TITLE_BAND of headroom above it for exactly
    // this (chart-shapes.ts). The allowance is that band, so a title band
    // shrunk below a label's height would fail here. Markers and connectors
    // have no such allowance and must stay inside `plot` itself.
    const labelRoom: Box = {
      ...plot,
      top: plot.top - TITLE_BAND,
      height: plot.height + TITLE_BAND,
    };
    for (const primitive of out) {
      expect(primitive.box.width).toBeGreaterThanOrEqual(0);
      expect(primitive.box.height).toBeGreaterThanOrEqual(0);
      const bounds = primitive.kind === "text" ? labelRoom : plot;
      expect(insideBox(primitive.box, bounds)).toBe(true);
    }
    // valueY is what the module scales every point through; sanity-check the
    // fixture actually spans the scale rather than sitting in one corner.
    expect(valueY(scale.max, scale, plot)).toBeCloseTo(plot.top, 6);
  });
});
