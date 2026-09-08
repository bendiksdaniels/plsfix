// Audit suite for the pure chart layout: the boxes PowerPoint would refuse
// (a negative width or height) and the primitives that would grow the group
// past the box it was placed in, over the inputs a real model produces - a
// negative value, a closing total below zero, more legend than the box holds.

import { describe, expect, it } from "vitest";
import { layoutChart, type Primitive, type Text } from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";

const base = {
  v: 1 as const,
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

// The box a chart of a normal Excel size lands in, and the smallest one
// chartSize can hand out (a 200 x 100 pt chart on the sheet).
const ROOMY: Box = { left: 100, top: 100, width: 480, height: 270 };
const TIGHT: Box = { left: 0, top: 0, width: 200, height: 120 };

function series(values: number[], name = "s"): ChartSeries {
  return {
    name,
    values,
    labels: values.map(String),
    colors: values.map(() => "#2EC4B6"),
  };
}

// A box the host would refuse outright, whatever it holds.
function refused(out: Primitive[]): Primitive[] {
  return out.filter((one) => one.box.width < 0 || one.box.height < 0);
}

// A primitive outside the chart's own box: the group PowerPoint builds is the
// union of its children, so one of these makes the link's geometry - and the
// corner every later refresh rebuilds at - drift away from the placement.
function outside(out: Primitive[], box: Box): Primitive[] {
  return out.filter(
    (one) =>
      one.box.left < box.left - 0.01 ||
      one.box.top < box.top - 0.01 ||
      one.box.left + one.box.width > box.left + box.width + 0.01 ||
      one.box.top + one.box.height > box.top + box.height + 0.01,
  );
}

function texts(out: Primitive[]): Text[] {
  return out.filter((one): one is Text => one.kind === "text");
}

describe("a value label the bar has no room for", () => {
  // One series, so there is no legend and the category band is all the space
  // under the plot: the label of the bar that reaches the bottom of the scale
  // has nowhere but that band to sit in.
  const negative: ChartData = {
    ...base,
    kind: "column",
    title: "Revenue",
    categories: ["A", "B", "C"],
    series: [series([100, -50, 200])],
  };

  it("keeps a negative bar's label inside the chart box", () => {
    const out = layoutChart(negative, ROOMY);
    expect(outside(out, ROOMY)).toEqual([]);
    const label = texts(out).find((one) => one.name === "label 0.1")!;
    const bar = out.find((one) => one.name === "bar 0.1")!;
    // Still below the bar it belongs to, which is what makes it readable.
    expect(label.box.top).toBeGreaterThanOrEqual(
      bar.box.top + bar.box.height - 0.01,
    );
  });

  it("keeps a waterfall's fall label inside the box too", () => {
    const bridge: ChartData = {
      ...base,
      kind: "waterfall",
      title: "Bridge",
      categories: ["Open", "Price", "Cost", "Close"],
      series: [series([100, 20, -110, 10])],
    };
    expect(outside(layoutChart(bridge, ROOMY), ROOMY)).toEqual([]);
  });
});

describe("a waterfall that closes below zero", () => {
  const bridge: ChartData = {
    ...base,
    kind: "waterfall",
    title: "Bridge",
    categories: ["Open", "Cost", "Close"],
    series: [series([100, -150, -50])],
  };

  it("draws the closing total as a bar the host will take", () => {
    const out = layoutChart(bridge, ROOMY);
    expect(refused(out)).toEqual([]);
    expect(outside(out, ROOMY)).toEqual([]);
    const baseline = out.find((one) => one.name === "baseline")!;
    const close = out.find((one) => one.name === "bar 0.2")!;
    // The closing total hangs below the zero line, like a negative column.
    expect(close.box.top).toBeCloseTo(baseline.box.top, 5);
    expect(close.box.height).toBeGreaterThan(0);
  });

  it("labels it below the bar, where a negative value's label belongs", () => {
    const out = layoutChart(bridge, ROOMY);
    const label = texts(out).find((one) => one.name === "label 0.2")!;
    const close = out.find((one) => one.name === "bar 0.2")!;
    expect(label.box.top).toBeCloseTo(close.box.top + close.box.height, 5);
  });
});

describe("a box too small for everything the chart wants", () => {
  const sixSeries: ChartData = {
    ...base,
    kind: "column",
    title: "Revenue by segment",
    categories: ["2024A", "2025E", "2026E"],
    series: [0, 1, 2, 3, 4, 5].map((j) =>
      series([10 + j, 20 + j, 30 + j], `Segment number ${String(j)}`),
    ),
  };
  const wideLegendPie: ChartData = {
    ...base,
    kind: "pie",
    title: "Mix",
    categories: Array.from(
      { length: 12 },
      (_unused, i) => `Category number ${String(i)}`,
    ),
    series: [series(Array.from({ length: 12 }, () => 10))],
  };

  it("still lays a six-series column out at sizes the host accepts", () => {
    const out = layoutChart(sixSeries, TIGHT);
    expect(refused(out)).toEqual([]);
    expect(outside(out, TIGHT)).toEqual([]);
    // The plot is not squeezed away: the bars are still drawn.
    const bars = out.filter((one) => one.name.startsWith("bar "));
    expect(bars).toHaveLength(18);
    expect(bars.every((one) => one.box.height > 0)).toBe(true);
  });

  it("cuts the legend to the rows the box has room for", () => {
    const drawn = texts(layoutChart(sixSeries, TIGHT)).filter((one) =>
      one.name.startsWith("legend "),
    );
    // Six one-per-row entries want 96 pt; half of the 102 pt under the title
    // is 51, so three rows are drawn and the chart keeps its plot.
    expect(drawn.map((one) => one.text)).toEqual([
      "Segment number 0",
      "Segment number 1",
      "Segment number 2",
    ]);
    expect(layoutChart(sixSeries, ROOMY).length).toBeGreaterThan(
      layoutChart(sixSeries, TIGHT).length,
    );
  });

  it("never gives a pie a negative radius", () => {
    const out = layoutChart(wideLegendPie, TIGHT);
    expect(refused(out)).toEqual([]);
    expect(outside(out, TIGHT)).toEqual([]);
    expect(out.filter((one) => one.kind === "wedge")).toHaveLength(12);
  });
});

describe("the inputs a modeller can still hand a chart", () => {
  it("keeps a stacked column with negative segments inside the plot", () => {
    const stacked: ChartData = {
      ...base,
      kind: "stackedColumn",
      title: "Movements",
      categories: ["A", "B"],
      series: [series([60, -40], "up"), series([-20, 30], "down")],
    };
    const out = layoutChart(stacked, ROOMY);
    expect(refused(out)).toEqual([]);
    expect(outside(out, ROOMY)).toEqual([]);
    // Two negative segments, each hanging below the baseline.
    const baseline = out.find((one) => one.name === "baseline")!;
    const below = out
      .filter((one) => one.name.startsWith("bar "))
      .filter((one) => one.box.top >= baseline.box.top - 0.01);
    expect(below).toHaveLength(2);
  });

  it("draws a two-point line with one connector and both labels inside", () => {
    const line: ChartData = {
      ...base,
      kind: "line",
      title: "Trend",
      categories: ["2024A", "2025E"],
      series: [series([120, 150])],
    };
    const out = layoutChart(line, ROOMY);
    expect(refused(out)).toEqual([]);
    expect(outside(out, ROOMY)).toEqual([]);
    expect(out.filter((one) => one.name.startsWith("line "))).toHaveLength(1);
    expect(out.filter((one) => one.name.startsWith("marker "))).toHaveLength(2);
  });

  it("gives a slice too thin to see two different angles, never a full circle", () => {
    const lopsided: ChartData = {
      ...base,
      kind: "pie",
      title: "Mix",
      categories: ["Rest", "Sliver"],
      series: [series([1_000_000, 1])],
    };
    const wedges = layoutChart(lopsided, ROOMY).filter(
      (one) => one.kind === "wedge",
    );
    expect(wedges).toHaveLength(2);
    // PowerPoint reads a start equal to its end as a full sweep, so the two
    // angles of even the thinnest slice have to stay apart.
    for (const wedge of wedges) {
      if (wedge.kind !== "wedge") continue;
      expect(wedge.start).not.toBe(wedge.end);
    }
  });

  it("draws nothing but the legend for a pie with no positive value", () => {
    const empty: ChartData = {
      ...base,
      kind: "pie",
      title: "Mix",
      categories: ["A", "B"],
      series: [series([0, 0])],
    };
    const out = layoutChart(empty, ROOMY);
    expect(out.filter((one) => one.kind === "wedge")).toEqual([]);
    expect(refused(out)).toEqual([]);
  });
});
