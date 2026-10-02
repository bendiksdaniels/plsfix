// The shared primitive builders, at the one place they can put a shape
// outside the chart: a horizontal bar's value label, which is placed past the
// end of the bar. A long number on a bar that reaches the plot's edge - forty
// nine-digit rows is what a real sensitivity table looks like - used to grow
// the group past the box it was placed in, which moves the link's geometry.

import { describe, expect, it } from "vitest";
import {
  layoutChart,
  LABEL_HEIGHT,
  type Primitive,
  type Text,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartKind, ChartSeries } from "./link/chart-model";

const base = {
  v: 1 as const,
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
  title: "Sensitivity",
};

// The box a chart of a normal Excel size lands in, and the smallest one
// chartSize hands out (a 200 x 100 pt chart on the sheet).
const ROOMY: Box = { left: 100, top: 100, width: 480, height: 270 };
const TIGHT: Box = { left: 0, top: 0, width: 200, height: 120 };

// A primitive outside the chart's own box: the group PowerPoint builds is the
// union of its children, so one of these drags the link's geometry - and the
// corner every later refresh rebuilds at - away from the placement.
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

function labelled(out: Primitive[], name: string): Text {
  return texts(out).find((one) => one.name === name)!;
}

function series(values: number[], name = "s", color = "#2EC4B6"): ChartSeries {
  return {
    name,
    values,
    labels: values.map(String),
    colors: values.map(() => color),
  };
}

// Forty rows of nine-digit values: the widest label a modeller's own number
// format produces, on the chart kind whose labels sit beside the bar.
function wide(kind: ChartKind, sign = 1): ChartData {
  const values = Array.from(
    { length: 40 },
    (_, i) => sign * (100000000 + i * 1000),
  );
  return {
    ...base,
    kind,
    categories: values.map((_, i) => `Row ${String(i)}`),
    series: [series(values)],
  };
}

describe("a bar chart's value labels", () => {
  it("keeps every primitive of forty nine-digit rows inside the chart box", () => {
    expect(outside(layoutChart(wide("bar"), ROOMY), ROOMY)).toEqual([]);
    expect(outside(layoutChart(wide("bar"), TIGHT), TIGHT)).toEqual([]);
  });

  it("keeps the negative side inside it too, label column and all", () => {
    expect(outside(layoutChart(wide("bar", -1), ROOMY), ROOMY)).toEqual([]);
    expect(outside(layoutChart(wide("bar", -1), TIGHT), TIGHT)).toEqual([]);
  });

  it("puts the label of the longest bar inside its end, reading right", () => {
    const out = layoutChart(wide("bar"), ROOMY);
    // Row 39 holds the largest value, so its bar reaches the plot's right
    // edge - which is the chart box's right edge - and its label has nowhere
    // outside the bar to go.
    const bar = out.find((one) => one.name === "bar 0.39")!;
    const label = labelled(out, "label 0.39");
    expect(bar.box.left + bar.box.width).toBeCloseTo(
      ROOMY.left + ROOMY.width,
      5,
    );
    expect(label.box.left + label.box.width).toBeCloseTo(
      bar.box.left + bar.box.width,
      5,
    );
    expect(label.align).toBe("r");
    expect(label.text).toBe("100039000");
  });

  it("leaves a label with room past the bar end where it has always been", () => {
    const short: ChartData = {
      ...base,
      kind: "bar",
      categories: ["A", "B"],
      series: [series([10, 4])],
    };
    const out = layoutChart(short, ROOMY);
    const bar = out.find((one) => one.name === "bar 0.1")!;
    const label = labelled(out, "label 0.1");
    expect(label.box.left).toBeCloseTo(bar.box.left + bar.box.width, 5);
    expect(label.align).toBe("l");
  });

  it("keeps a negative bar's label on its far side, in the label column", () => {
    const both: ChartData = {
      ...base,
      kind: "bar",
      categories: ["Down", "Up"],
      series: [series([-20, 25])],
    };
    const out = layoutChart(both, ROOMY);
    const bar = out.find((one) => one.name === "bar 0.0")!;
    const label = labelled(out, "label 0.0");
    expect(label.box.left + label.box.width).toBeCloseTo(bar.box.left, 5);
    expect(label.align).toBe("r");
    expect(outside(out, ROOMY)).toEqual([]);
  });

  it("reads white when the clamp puts it on a dark bar", () => {
    // The brand's first series colour is the navy the pane's own ink nearly
    // is: charcoal printed inside that bar is unreadable, so a clamped label
    // takes the white a dark stacked segment's label has always taken.
    const navy: ChartData = {
      ...base,
      kind: "bar",
      categories: ["A", "B"],
      series: [series([10, 25], "s", "#14213D")],
    };
    const out = layoutChart(navy, ROOMY);
    const inside = labelled(out, "label 0.1");
    const outside = labelled(out, "label 0.0");
    expect(inside.align).toBe("r");
    expect(inside.color).toBe("#FFFFFF");
    // The short bar's label still sits past its end, on the slide's ground.
    expect(outside.align).toBe("l");
    expect(outside.color).toBe(base.ink);
  });

  it("keeps the ink when the clamp puts it on a light bar", () => {
    const mint: ChartData = {
      ...base,
      kind: "bar",
      categories: ["A", "B"],
      series: [series([10, 25], "s", "#2EC4B6")],
    };
    const label = labelled(layoutChart(mint, ROOMY), "label 0.1");
    expect(label.align).toBe("r");
    expect(label.color).toBe(base.ink);
  });

  it("gives every label the same height it always had", () => {
    const out = layoutChart(wide("bar"), ROOMY);
    const heights = texts(out)
      .filter((one) => one.name.startsWith("label "))
      .map((one) => one.box.height);
    expect(new Set(heights)).toEqual(new Set([LABEL_HEIGHT]));
  });
});

describe("the kinds the clamp does not touch", () => {
  it("keeps every nine-digit column bar inside the box and draws no value label so dense a chart cannot hold", () => {
    const out = layoutChart(wide("column"), ROOMY);
    expect(outside(out, ROOMY)).toEqual([]);
    const bar = out.find((one) => one.name === "bar 0.39");
    expect(bar).toBeDefined();
    // A 12 pt slot cannot hold a nine-digit value label without wrapping it:
    // the layout draws none for the whole chart rather than stack every one
    // of them into a column of digits (dense-chart mitigation, chart-place).
    expect(texts(out).some((one) => one.name.startsWith("label "))).toBe(false);
  });

  it("still centres a column's label over its own bar, unclamped, once the slot has room", () => {
    const roomy: ChartData = {
      ...base,
      kind: "column",
      categories: ["Row 0", "Row 1"],
      series: [series([100000000, 100000000])],
    };
    const out = layoutChart(roomy, ROOMY);
    const bar = out.find((one) => one.name === "bar 0.0")!;
    const label = labelled(out, "label 0.0");
    expect(label.box.left).toBeCloseTo(bar.box.left, 5);
    expect(label.box.width).toBeCloseTo(bar.box.width, 5);
    expect(label.align).toBe("c");
  });

  it("centres a stacked bar's label inside the segment, as it always did", () => {
    const stacked: ChartData = {
      ...base,
      kind: "stackedBar",
      categories: ["A", "B"],
      series: [series([100000000, 100000000], "most"), series([1, 1], "rest")],
    };
    const out = layoutChart(stacked, ROOMY);
    const bar = out.find((one) => one.name === "bar 0.0")!;
    const label = labelled(out, "label 0.0");
    expect(label.box).toEqual(bar.box);
    expect(label.align).toBe("c");
    expect(outside(out, ROOMY)).toEqual([]);
  });

  // A stacked bar's segments share their row the way a stacked column's share
  // their slot: same band, same thickness, the stacking done by the values.
  // The clustered offset put each series in a band of its own, which pushed
  // the outer ones past the chart box as soon as there were three.
  it.each([2, 3, 6])("bands %i stacked series into one row each", (count) => {
    const stacked: ChartData = {
      ...base,
      kind: "stackedBar",
      categories: ["A", "B"],
      series: Array.from({ length: count }, (_, j) =>
        series([100, 100], `s${String(j)}`),
      ),
    };
    const out = layoutChart(stacked, ROOMY);
    expect(outside(out, ROOMY)).toEqual([]);
    const row = out.filter(
      (one) => one.name.startsWith("bar ") && one.name.endsWith(".0"),
    );
    expect(row).toHaveLength(count);
    expect(new Set(row.map((one) => one.box.top)).size).toBe(1);
    expect(new Set(row.map((one) => one.box.height)).size).toBe(1);
  });

  it("keeps a waterfall's labels over and under its bars", () => {
    const bridge: ChartData = {
      ...base,
      kind: "waterfall",
      categories: ["Open", "Price", "Cost", "Close"],
      series: [series([100000000, 20000000, -110000000, 10000000])],
    };
    const out = layoutChart(bridge, ROOMY);
    expect(outside(out, ROOMY)).toEqual([]);
    const bar = out.find((one) => one.name === "bar 0.0")!;
    const label = labelled(out, "label 0.0");
    expect(label.box.width).toBeCloseTo(bar.box.width, 5);
    expect(label.align).toBe("c");
  });
});
