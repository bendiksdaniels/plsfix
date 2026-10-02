// Direct proof of chart-shapes-parts.ts's placement-at-scale behaviour:
// legend draws every swatch and label inside its band and never past the
// rows the band's height affords, pieWedges gives every wedge the same
// centred square box with angles closing a full turn (within 1e-6) and
// normalised to (-180, 180], negatives skipped like zeros, and the
// 12-slice cap. Then layoutChart (chart-shapes.ts) at the edges
// chart-shapes.test.ts leaves out: a single category, a zero-value bar, and
// the 40-point/6-series column and 40-point waterfall caps, every primitive
// non-negative and inside the box. The smaller value/placement primitives
// (lineBetween, truncate, bridgeRange, segmentLabel, ...) live in the
// sibling chart-shapes-parts.test.ts to keep both files under the line cap.
import { describe, expect, it } from "vitest";
import { legend, pieWedges } from "./chart-shapes-parts";
import {
  LEGEND_BAND,
  layoutChart,
  type Ellipse,
  type Rect,
  type Wedge,
} from "./chart-shapes";
import type { Box } from "./layout";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  type ChartData,
} from "./link/chart-model";

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

describe("legend", () => {
  const pie: ChartData = {
    v: 1,
    kind: "pie",
    title: null,
    categories: ["Alpha", "Beta", "Gamma", "Delta"],
    series: [
      {
        name: "s",
        values: [1, 1, 1, 1],
        labels: ["1", "1", "1", "1"],
        colors: ["#111111", "#222222", "#333333", "#444444"],
      },
    ],
    font: "Arial",
    ink: "#000000",
    titleColor: "#000000",
  };

  it("draws every swatch and label inside the band it was given, even once items wrap", () => {
    const band: Box = { left: 20, top: 200, width: 250, height: 48 };
    const out = legend(pie, band);
    expect(out.length).toBeGreaterThan(0);
    expect(out.every((primitive) => insideBox(primitive.box, band))).toBe(true);
  });

  it("never draws past the rows the band's own height affords, even if that drops entries", () => {
    // A width no single item shares a row under, and a height for one row
    // only: every category after the first must be left off entirely.
    const band: Box = { left: 0, top: 0, width: 1, height: LEGEND_BAND };
    const out = legend(pie, band);
    expect(out.filter((p) => p.name.startsWith("legend swatch"))).toHaveLength(
      1,
    );
    const labels = out.filter(
      (p) =>
        p.name.startsWith("legend ") && !p.name.startsWith("legend swatch"),
    );
    expect(labels).toHaveLength(1);
    expect(labels[0]?.kind === "text" ? labels[0].text : null).toBe("Alpha");
  });
});

function pieData(categories: string[], values: number[]): ChartData {
  return {
    v: 1,
    kind: "pie",
    title: null,
    categories,
    series: [
      {
        name: "s",
        values,
        labels: values.map(String),
        colors: categories.map(() => "#2EC4B6"),
      },
    ],
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
}

function wedgeSpan(wedge: Wedge): number {
  return (((wedge.end - wedge.start) % 360) + 360) % 360;
}

describe("pieWedges", () => {
  const plot: Box = { left: 100, top: 50, width: 200, height: 200 };

  it("gives every wedge the same square box, centred in the plot", () => {
    const wedges = pieWedges(pieData(["A", "B"], [1, 1]), plot).filter(
      (p): p is Wedge => p.kind === "wedge",
    );
    for (const wedge of wedges) {
      expect(wedge.box.width).toBe(wedge.box.height);
      expect(insideBox(wedge.box, plot)).toBe(true);
    }
    expect(new Set(wedges.map((w) => JSON.stringify(w.box))).size).toBe(1);
  });

  it("skips a negative slice like a zero one: no wedge, no share of the angle", () => {
    const data = pieData(["A", "B", "C"], [10, 20, -5]);
    const wedges = pieWedges(data, plot).filter(
      (p): p is Wedge => p.kind === "wedge",
    );
    expect(wedges).toHaveLength(2);
    const spans = wedges.map(wedgeSpan);
    expect(Math.abs(spans[0]! - 120)).toBeLessThan(1e-9); // 10 / (10+20) * 360
    expect(Math.abs(spans[1]! - 240)).toBeLessThan(1e-9); // 20 / (10+20) * 360
  });

  it("closes a full turn within 1e-6 and keeps every angle normalised to (-180, 180]", () => {
    const categories = Array.from(
      { length: CHART_MAX_SLICES },
      (_unused, i) => `S${String(i)}`,
    );
    const wedges = pieWedges(
      pieData(
        categories,
        categories.map(() => 1),
      ),
      plot,
    ).filter((p): p is Wedge => p.kind === "wedge");
    expect(wedges).toHaveLength(CHART_MAX_SLICES);
    const totalSpan = wedges.reduce((sum, w) => sum + wedgeSpan(w), 0);
    expect(Math.abs(totalSpan - 360)).toBeLessThan(1e-6);
    for (const wedge of wedges) {
      for (const angle of [wedge.start, wedge.end]) {
        expect(angle).toBeGreaterThan(-180);
        expect(angle).toBeLessThanOrEqual(180);
      }
    }
  });

  it("draws nothing at all once every value is zero or negative", () => {
    expect(pieWedges(pieData(["A", "B"], [0, -3]), plot)).toEqual([]);
  });

  it("draws a lone positive slice among a zero and a negative as an ellipse, not a wedge", () => {
    const out = pieWedges(pieData(["A", "B", "C"], [7, 0, -3]), plot);
    expect(out).toHaveLength(1);
    expect(out[0]?.kind).toBe("ellipse");
    const ellipse = out[0] as Ellipse;
    expect(ellipse.box.width).toBe(ellipse.box.height);
    expect(insideBox(ellipse.box, plot)).toBe(true);
  });
});

describe("layoutChart at the edges chart-shapes.test.ts leaves out", () => {
  const box: Box = { left: 0, top: 0, width: 800, height: 400 };
  const base = {
    v: 1 as const,
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
    title: null,
  };

  it("draws a single-category column with no crash and every primitive inside the box", () => {
    const data: ChartData = {
      ...base,
      kind: "column",
      categories: ["Only"],
      series: [
        { name: "s", values: [42], labels: ["42"], colors: ["#2EC4B6"] },
      ],
    };
    const out = layoutChart(data, box);
    const bars = out.filter(
      (p): p is Rect => p.kind === "rect" && p.name.startsWith("bar"),
    );
    expect(bars).toHaveLength(1);
    expect(out.every((p) => insideBox(p.box, box))).toBe(true);
  });

  it("gives a zero-value bar a zero-height box on the baseline, never a negative one", () => {
    const data: ChartData = {
      ...base,
      kind: "column",
      categories: ["A", "B"],
      series: [
        {
          name: "s",
          values: [0, 15],
          labels: ["0", "15"],
          colors: ["#2EC4B6", "#2EC4B6"],
        },
      ],
    };
    const out = layoutChart(data, box);
    const bars = out.filter(
      (p): p is Rect => p.kind === "rect" && p.name.startsWith("bar"),
    );
    const baseline = out.find((p) => p.name === "baseline")!;
    expect(bars[0]!.box.height).toBe(0);
    expect(bars[0]!.box.top).toBeCloseTo(baseline.box.top, 6);
    expect(out.every((p) => insideBox(p.box, box))).toBe(true);
  });

  it("keeps every primitive inside the box at the 40-point, 6-series cap", () => {
    const categories = Array.from(
      { length: CHART_MAX_POINTS },
      (_unused, i) => `C${String(i)}`,
    );
    const series = Array.from({ length: CHART_MAX_SERIES }, (_unused, j) => ({
      name: `s${String(j)}`,
      values: categories.map((_unused2, i) => (i % 2 === 0 ? i + j : -(i + j))),
      labels: categories.map(() => ""),
      colors: categories.map(() => "#2EC4B6"),
    }));
    const data: ChartData = { ...base, kind: "column", categories, series };
    const out = layoutChart(data, box);
    const bars = out.filter(
      (p): p is Rect => p.kind === "rect" && p.name.startsWith("bar"),
    );
    expect(bars).toHaveLength(CHART_MAX_POINTS * CHART_MAX_SERIES);
    expect(out.every((p) => insideBox(p.box, box))).toBe(true);
  });

  it("floats a 40-point bridge with 39 connectors, all inside the box", () => {
    const categories = Array.from(
      { length: CHART_MAX_POINTS },
      (_unused, i) => `C${String(i)}`,
    );
    const values = categories.map((_unused, i) =>
      i === 0 || i === categories.length - 1 ? 100 : i % 2 === 0 ? i : -i,
    );
    const data: ChartData = {
      ...base,
      kind: "waterfall",
      categories,
      series: [
        {
          name: "s",
          values,
          labels: values.map(String),
          colors: categories.map(() => "#2EC4B6"),
        },
      ],
    };
    const out = layoutChart(data, box);
    expect(out.filter((p) => p.name.startsWith("connector"))).toHaveLength(
      CHART_MAX_POINTS - 1,
    );
    expect(out.every((p) => insideBox(p.box, box))).toBe(true);
  });
});
