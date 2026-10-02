// Hunt pass 1: shape-built charts at value/count edges layoutChart's existing
// suites leave out - a single category, a line chart with a negative,
// all-zero or mixed-sign series, and 200 points past the 40-point cap. Pure
// geometry: must degrade to finite, in-box boxes, in scale order, never NaN.

import { describe, expect, it } from "vitest";
import {
  layoutChart,
  type Ellipse,
  type Line,
  type Primitive,
  type Text,
} from "../../src/chart-shapes";
import type { Box } from "../../src/layout";
import type { ChartData, ChartSeries } from "../../src/link/chart-model";

const base = {
  v: 1 as const,
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
  title: null,
};

const ROOMY: Box = { left: 100, top: 100, width: 480, height: 270 };

function series(values: number[], name = "s"): ChartSeries {
  return {
    name,
    values,
    labels: values.map(String),
    colors: values.map(() => "#2EC4B6"),
  };
}

function chart(
  kind: ChartData["kind"],
  categories: string[],
  values: number[],
): ChartData {
  return { ...base, kind, categories, series: [series(values)] };
}

// The same finiteness/containment check every existing chart-shapes suite
// runs: nothing NaN, nothing negative-sized (PowerPoint's own InvalidArgument
// shape), and nothing straying outside the box the group must stay inside.
function outside(out: Primitive[], box: Box): Primitive[] {
  return out.filter(
    (one) =>
      !Number.isFinite(one.box.left) ||
      !Number.isFinite(one.box.top) ||
      !Number.isFinite(one.box.width) ||
      !Number.isFinite(one.box.height) ||
      one.box.width < 0 ||
      one.box.height < 0 ||
      one.box.left < box.left - 0.01 ||
      one.box.top < box.top - 0.01 ||
      one.box.left + one.box.width > box.left + box.width + 0.01 ||
      one.box.top + one.box.height > box.top + box.height + 0.01,
  );
}

// A series' own point-to-point connectors, named "line <series>.<point>" by
// lineBetween - never the always-present zero-axis "baseline", which is the
// same primitive kind but not a connector.
function lines(out: Primitive[]): Line[] {
  return out.filter(
    (one): one is Line => one.kind === "line" && one.name.startsWith("line "),
  );
}
function markers(out: Primitive[]): Ellipse[] {
  return out.filter((one): one is Ellipse => one.kind === "ellipse");
}
function texts(out: Primitive[]): Text[] {
  return out.filter((one): one is Text => one.kind === "text");
}

describe("a single category (one point): every layoutChart test elsewhere uses at least two", () => {
  it.each(["column", "bar", "line"] as const)(
    "%s: one bar/marker, no connector, everything inside the box",
    (kind) => {
      const out = layoutChart(chart(kind, ["Only"], [42]), ROOMY);

      expect(outside(out, ROOMY)).toEqual([]);
      if (kind === "line") {
        expect(lines(out)).toHaveLength(0); // no point before it to connect
        expect(markers(out)).toHaveLength(1);
      } else {
        expect(out.filter((p) => p.name.startsWith("bar "))).toHaveLength(1);
      }
      expect(
        texts(out).some((t) => t.name.startsWith("label ") && t.text === "42"),
      ).toBe(true);
    },
  );
});

describe("a line chart whose every value is negative", () => {
  it("places every marker inside the plot, still one connector fewer than points, all rising flagged correctly", () => {
    const data = chart("line", ["A", "B", "C"], [-10, -40, -20]);

    const out = layoutChart(data, ROOMY);

    expect(outside(out, ROOMY)).toEqual([]);
    expect(markers(out)).toHaveLength(3);
    const segments = lines(out);
    expect(segments).toHaveLength(2);
    // -10 -> -40 falls, -40 -> -20 climbs back up.
    expect(segments.map((l) => l.rising)).toEqual([false, true]);
    expect(texts(out).map((t) => t.text)).toEqual(
      expect.arrayContaining(["-10", "-40", "-20"]),
    );
  });
});

describe("a line chart whose values cross zero (mixed sign)", () => {
  it("keeps the whole series inside the plot on both sides of the baseline", () => {
    const data = chart("line", ["A", "B", "C", "D"], [50, -30, 0, 20]);

    const out = layoutChart(data, ROOMY);

    expect(outside(out, ROOMY)).toEqual([]);
    expect(markers(out)).toHaveLength(4);
    expect(lines(out)).toHaveLength(3);
    // The highest point (50) and the lowest (-30) both clamp to their edge of
    // the plot rather than spill past it - the one thing a shared min/max
    // scale has to get right for a series that is not all one sign.
    const plotTop = Math.min(...markers(out).map((m) => m.box.top));
    const plotBottom = Math.max(
      ...markers(out).map((m) => m.box.top + m.box.height),
    );
    expect(plotTop).toBeGreaterThanOrEqual(ROOMY.top - 0.01);
    expect(plotBottom).toBeLessThanOrEqual(ROOMY.top + ROOMY.height + 0.01);

    // outside() only proves every marker stays in the box; it says nothing
    // about them landing in the RIGHT place relative to one another. A
    // shared scale across a mixed-sign series is what this pins: by value,
    // 50 sits highest on the plot (smallest top), then 20, then 0, then -30.
    function topOf(pointIndex: number): number {
      return markers(out).find(
        (marker) => marker.name === `marker 0.${String(pointIndex)}`,
      )!.box.top;
    }
    const top50 = topOf(0);
    const top20 = topOf(3);
    const top0 = topOf(2);
    const topMinus30 = topOf(1);
    expect(top50).toBeLessThan(top20);
    expect(top20).toBeLessThan(top0);
    expect(top0).toBeLessThan(topMinus30);
  });
});

describe("a line chart whose every value is zero", () => {
  it("draws every marker on one flat baseline, no crash on a zero-width scale range", () => {
    const data = chart("line", ["A", "B", "C"], [0, 0, 0]);

    const out = layoutChart(data, ROOMY);

    expect(outside(out, ROOMY)).toEqual([]);
    const tops = markers(out).map((m) => m.box.top);
    expect(new Set(tops.map((t) => Math.round(t * 100)))).toHaveProperty(
      "size",
      1,
    );
  });
});

describe("far past the real 40-point cap (layoutChart itself has no ceiling)", () => {
  it.each(["column", "line"] as const)(
    "%s: 200 points still lay out finite and in-box, the cap is the model validator's job, not this one's",
    (kind) => {
      const categories = Array.from({ length: 200 }, (_, i) => String(i));
      const values = categories.map((_, i) => (i % 2 === 0 ? i : -i));
      const data = chart(kind, categories, values);

      const out = layoutChart(data, ROOMY);

      expect(outside(out, ROOMY)).toEqual([]);
      expect(out.length).toBeGreaterThan(0);
    },
  );
});
