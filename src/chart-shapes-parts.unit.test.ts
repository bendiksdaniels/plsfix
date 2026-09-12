// Direct proof of chart-shapes-parts.ts's smaller value and placement
// primitives: lineBetween's normalisation, truncate's ellipsis rule, the
// bar/stack running-cursor maths, bridgeRange's below flag, rowOrder's
// overlap ranking, segmentLabel's four placement branches, legendItems and
// packLegendRows' wrap rule. Legend drawing, pie wedges and the
// layoutChart-at-scale proofs live in the sibling
// chart-shapes-parts.scale.test.ts to keep both files under the line cap.
import { describe, expect, it } from "vitest";
import type { BridgeSeries } from "./chartmath";
import {
  barRange,
  bridgeRange,
  legendItems,
  lineBetween,
  packLegendRows,
  rowOrder,
  segmentLabel,
  stackSegment,
  truncate,
} from "./chart-shapes-parts";
import { LABEL_PAD, LABEL_SIZE, SWATCH, textWidth } from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";

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

describe("lineBetween", () => {
  it("normalises the box to a non-negative width and height whichever end is given first", () => {
    const forward = lineBetween(10, 50, 30, 20, "#000000", 1, "l");
    expect(forward.box).toEqual({ left: 10, top: 20, width: 20, height: 30 });
    const backward = lineBetween(30, 20, 10, 50, "#000000", 1, "l");
    expect(backward.box).toEqual(forward.box);
  });

  it("flags rising only for a segment that actually climbs left to right", () => {
    expect(lineBetween(10, 50, 30, 20, "#000000", 1, "l").rising).toBe(true); // y falls, chart rises
    expect(lineBetween(30, 20, 10, 50, "#000000", 1, "l").rising).toBe(false);
    // A flat segment (equal y) is not a climb either: strictly less-than, not
    // less-than-or-equal.
    expect(lineBetween(10, 40, 30, 40, "#000000", 1, "l").rising).toBe(false);
  });
});

describe("truncate", () => {
  it("leaves text alone once it already fits", () => {
    expect(truncate("Q3", LABEL_SIZE, 1_000)).toBe("Q3");
  });

  it("cuts to the widest prefix an ellipsis still fits after, never past zero characters", () => {
    const text = "Quarterly revenue by segment";
    const full = textWidth(text, LABEL_SIZE);
    const cut = truncate(text, LABEL_SIZE, full / 2);
    expect(cut.length).toBeLessThan(text.length);
    expect(cut.endsWith("…")).toBe(true);
    expect(textWidth(cut, LABEL_SIZE)).toBeLessThanOrEqual(full / 2 + 0.01);
  });

  it("falls back to a bare ellipsis rather than an empty string when nothing fits", () => {
    expect(truncate("Quarterly revenue", LABEL_SIZE, 0)).toBe("…");
  });
});

describe("stackSegment and barRange", () => {
  it("stacks positive segments upward from zero and negative ones downward, independently", () => {
    const cursor = { pos: 0, neg: 0 };
    expect(stackSegment(cursor, 5)).toEqual([0, 5]);
    expect(stackSegment(cursor, 3)).toEqual([5, 8]);
    expect(stackSegment(cursor, -2)).toEqual([-2, 0]);
    expect(stackSegment(cursor, -4)).toEqual([-6, -2]);
    expect(cursor).toEqual({ pos: 8, neg: -6 });
  });

  it("un-stacked, spans from zero to the value on whichever side it falls, zero included", () => {
    const cursor = { pos: 0, neg: 0 };
    expect(barRange(false, cursor, 12)).toEqual([0, 12]);
    expect(barRange(false, cursor, -12)).toEqual([-12, 0]);
    expect(barRange(false, cursor, 0)).toEqual([0, 0]);
    // Unstacked bars never touch the cursor at all.
    expect(cursor).toEqual({ pos: 0, neg: 0 });
  });

  it("stacked, threads the same running cursor stackSegment would", () => {
    const cursor = { pos: 0, neg: 0 };
    expect(barRange(true, cursor, 6)).toEqual([0, 6]);
    expect(barRange(true, cursor, 4)).toEqual([6, 10]);
  });
});

describe("bridgeRange", () => {
  function bridge(base: number, rise: number, fall: number): BridgeSeries {
    return { base: [base], rise: [rise], fall: [fall] };
  }

  it("puts a rising step's label above: low to high, never flagged below", () => {
    expect(bridgeRange(bridge(10, 5, 0), 0)).toEqual({
      lo: 10,
      hi: 15,
      below: false,
    });
  });

  it("always flags a genuine fall below, magnitude sitting on top of where it lands", () => {
    // A -12 delta landing at level 8: base = 8, fall = 12, to = 20.
    expect(bridgeRange(bridge(8, 0, 12), 0)).toEqual({
      lo: 8,
      hi: 20,
      below: true,
    });
  });

  it("flags below for a negative opening or closing total too, even with no fall at all", () => {
    // An opening/closing total is carried as base 0, rise = the total itself;
    // a negative total still has to read low-to-high with its label below.
    expect(bridgeRange(bridge(0, -5, 0), 0)).toEqual({
      lo: -5,
      hi: 0,
      below: true,
    });
  });
});

describe("rowOrder", () => {
  function series(values: number[]): ChartSeries {
    return {
      name: "s",
      values,
      labels: values.map(String),
      colors: values.map(() => "#000000"),
    };
  }

  it("keeps input order when nothing overlaps", () => {
    const rows = [series([-5, 0]), series([20, 0]), series([8, 0])];
    expect(rowOrder(rows, 0, false)).toEqual([0, 1, 2]);
  });

  it("orders overlapping rows by the widest swing first, sign ignored", () => {
    const rows = [series([-5, 0]), series([20, 0]), series([8, 0])];
    expect(rowOrder(rows, 0, true)).toEqual([1, 2, 0]);
  });
});

describe("segmentLabel", () => {
  const data: ChartData = {
    v: 1,
    kind: "column",
    title: null,
    categories: ["A"],
    series: [],
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
  const box: Box = { left: 100, top: 200, width: 40, height: 60 };
  // A chart box far larger than any label, so the clamp that keeps a bar's
  // label inside the chart (chart-shapes-parts.test.ts covers it) never fires.
  const chart: Box = { left: 0, top: 0, width: 2000, height: 2000 };
  const one = (value: number, color = "#2EC4B6"): ChartSeries => ({
    name: "s",
    values: [value],
    labels: [String(value)],
    colors: [color],
  });

  it("drops a stacked segment's label once its own box is thinner than the minimum", () => {
    const out = segmentLabel(
      data,
      one(1),
      0,
      0,
      { ...box, height: 2 },
      true,
      true,
      chart,
    );
    expect(out).toEqual([]);
  });

  it("centres a stacked segment's label, white on a dark fill and ink on a light one", () => {
    const [onDark] = segmentLabel(
      data,
      one(1, "#14213D"),
      0,
      0,
      box,
      true,
      true,
      chart,
    );
    expect(onDark?.color).toBe("#FFFFFF");
    expect(onDark?.align).toBe("c");
    expect(insideBox(onDark!.box, box)).toBe(true);

    const [onLight] = segmentLabel(data, one(1), 0, 0, box, true, true, chart);
    expect(onLight?.color).toBe(data.ink);
  });

  it("sits a vertical (column) label above a positive bar and below a negative one", () => {
    const [above] = segmentLabel(data, one(10), 0, 0, box, false, true, chart);
    expect(above?.box.top).toBe(box.top - above!.box.height);

    const [below] = segmentLabel(data, one(-10), 0, 0, box, false, true, chart);
    expect(below?.box.top).toBe(box.top + box.height);
  });

  it("sits a horizontal (bar) label to the right of a positive bar and left of a negative one", () => {
    const [right] = segmentLabel(data, one(10), 0, 0, box, false, false, chart);
    expect(right?.box.left).toBe(box.left + box.width);
    expect(right?.align).toBe("l");

    const [left] = segmentLabel(data, one(-10), 0, 0, box, false, false, chart);
    expect(left?.align).toBe("r");
    expect(left!.box.left + left!.box.width).toBeCloseTo(box.left, 6);
  });
});

describe("legendItems", () => {
  it("reads a pie's own categories, and every other kind's series names", () => {
    const base = {
      v: 1 as const,
      title: null,
      font: "Arial",
      ink: "#000000",
      titleColor: "#000000",
    };
    const pie: ChartData = {
      ...base,
      kind: "pie",
      categories: ["A", "B"],
      series: [
        {
          name: "ignored",
          values: [1, 2],
          labels: ["1", "2"],
          colors: ["#000", "#111"],
        },
      ],
    };
    expect(legendItems(pie)).toEqual(["A", "B"]);

    const column: ChartData = {
      ...base,
      kind: "column",
      categories: ["x"],
      series: [
        { name: "Revenue", values: [1], labels: ["1"], colors: ["#000"] },
        { name: "Cost", values: [1], labels: ["1"], colors: ["#111"] },
      ],
    };
    expect(legendItems(column)).toEqual(["Revenue", "Cost"]);
  });
});

function itemWidth(text: string): number {
  return SWATCH + LABEL_PAD + textWidth(text, LABEL_SIZE) + LABEL_PAD;
}

describe("packLegendRows", () => {
  it("keeps one empty row rather than none, so an empty legend can still be grouped", () => {
    expect(packLegendRows([], 500)).toEqual([[]]);
  });

  it("packs items that fit together onto one row", () => {
    const width = itemWidth("A") + itemWidth("BB");
    expect(packLegendRows(["A", "BB"], width)).toEqual([[0, 1]]);
  });

  it("wraps once the next item would overflow the row, one item at a time if it must", () => {
    const width = itemWidth("A") + itemWidth("BB") - 1;
    expect(packLegendRows(["A", "BB", "CCC"], width)).toEqual([[0], [1], [2]]);
  });

  it("keeps a single item too wide for the row rather than dropping it", () => {
    expect(packLegendRows(["A whole legend entry on its own"], 1)).toEqual([
      [0],
    ]);
  });
});
