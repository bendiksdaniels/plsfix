import { describe, expect, it } from "vitest";
import type { ChartData } from "./link/chart-model";
import {
  chartSize,
  layoutChart,
  type Ellipse,
  type Line,
  type Primitive,
  type Rect,
  type Text,
  type Wedge,
} from "./chart-shapes";

const box = { left: 100, top: 100, width: 480, height: 270 };
const base = {
  v: 1 as const,
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
  title: "Revenue",
};
const rects = (p: Primitive[]) =>
  p.filter((x): x is Rect => x.kind === "rect" && x.name.startsWith("bar"));
const texts = (p: Primitive[]) => p.filter((x): x is Text => x.kind === "text");
const lineShapes = (p: Primitive[]) =>
  p.filter((x): x is Line => x.kind === "line");
// A box PowerPoint would refuse (InvalidArgument on a negative width or
// height) can still land inside these outer bounds by coordinate alone, so a
// negative side must fail this check on its own before the position checks.
const inside = (b: {
  left: number;
  top: number;
  width: number;
  height: number;
}) =>
  b.width >= 0 &&
  b.height >= 0 &&
  b.left >= box.left - 0.01 &&
  b.top >= box.top - 0.01 &&
  b.left + b.width <= box.left + box.width + 0.01 &&
  b.top + b.height <= box.top + box.height + 0.01;

const column: ChartData = {
  ...base,
  kind: "column",
  categories: ["A", "B", "C"],
  series: [
    {
      name: "s",
      values: [100, -50, 200],
      labels: ["100", "(50)", "200"],
      colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    },
  ],
};

describe("layoutChart column", () => {
  const out = layoutChart(column, box);
  it("draws one bar per category inside the box, tallest for the largest value", () => {
    const bars = rects(out);
    expect(bars).toHaveLength(3);
    expect(bars.every((b) => inside(b.box))).toBe(true);
    expect(bars[2]!.box.height).toBeGreaterThan(bars[0]!.box.height);
  });
  it("hangs a negative bar below the baseline", () => {
    const [a, b] = rects(out);
    const baseline = out.find(
      (x) => x.kind === "line" && x.name === "baseline",
    )!;
    expect(a!.box.top + a!.box.height).toBeCloseTo(baseline.box.top, 0);
    expect(b!.box.top).toBeCloseTo(baseline.box.top, 0);
  });
  it("labels every bar with its text, above a positive and below a negative bar, and names the categories", () => {
    const labels = texts(out).filter((t) => t.name.startsWith("label"));
    expect(labels.map((t) => t.text)).toEqual(["100", "(50)", "200"]);
    const [a, b] = rects(out);
    expect(labels[0]!.box.top + labels[0]!.box.height).toBeLessThanOrEqual(
      a!.box.top + 0.01,
    );
    expect(labels[1]!.box.top).toBeGreaterThanOrEqual(
      b!.box.top + b!.box.height - 0.01,
    );
    expect(
      texts(out)
        .filter((t) => t.name.startsWith("category"))
        .map((t) => t.text),
    ).toEqual(["A", "B", "C"]);
  });
  it("titles the chart in the title colour and draws no legend for one series", () => {
    const title = texts(out).find((t) => t.name === "title")!;
    expect(title).toMatchObject({
      text: "Revenue",
      bold: true,
      color: "#14213D",
      size: 12,
    });
    expect(out.some((x) => x.name.startsWith("legend"))).toBe(false);
  });
  it("gives every primitive a unique name", () => {
    expect(new Set(out.map((x) => x.name)).size).toBe(out.length);
  });
});

describe("layoutChart stacked and bars", () => {
  const stacked: ChartData = {
    ...column,
    kind: "stackedColumn",
    series: [
      {
        name: "a",
        values: [60, 40, 80],
        labels: ["60", "40", "80"],
        colors: ["#14213D", "#14213D", "#14213D"],
      },
      {
        name: "b",
        values: [40, 60, 20],
        labels: ["40", "60", "20"],
        colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
      },
    ],
  };
  it("stacks two series into one column per category with white labels inside the dark segments", () => {
    const out = layoutChart(stacked, box);
    const bars = rects(out);
    expect(bars).toHaveLength(6);
    expect(bars[0]!.box.left).toBeCloseTo(bars[3]!.box.left);
    const inA = texts(out).find((t) => t.name === "label 0.0")!;
    expect(inA.color).toBe("#FFFFFF");
    expect(
      out.filter((x) => x.name.startsWith("legend")).length,
    ).toBeGreaterThan(0);
  });
  it("turns a bar chart on its side: bars grow to the right from a vertical baseline", () => {
    const out = layoutChart({ ...column, kind: "bar" }, box);
    const [a, , c] = rects(out);
    expect(c!.box.width).toBeGreaterThan(a!.box.width);
    expect(a!.box.top).toBeLessThan(c!.box.top);
    const baseline = out.find((x) => x.name === "baseline")!;
    expect(baseline.box.width).toBe(0);
  });
  it("draws a tornado row with the longer bar first so the shorter stays visible", () => {
    const tornado: ChartData = {
      ...column,
      kind: "bar",
      overlap: true,
      categories: ["Volume", "Price"],
      series: [
        {
          name: "Low",
          values: [-20, -5],
          labels: ["(20)", "(5)"],
          colors: ["#B00020", "#B00020"],
        },
        {
          name: "High",
          values: [15, 25],
          labels: ["15", "25"],
          colors: ["#2EC4B6", "#2EC4B6"],
        },
      ],
    };
    const out = layoutChart(tornado, box);
    const row = rects(out).filter((r) => r.name.endsWith(".1"));
    expect(row).toHaveLength(2);
    expect(row[0]!.box.width).toBeGreaterThan(row[1]!.box.width);
  });
});

describe("layoutChart line", () => {
  // Series 0 (100, -50, 200) is column's own; its last point sits at the
  // scale max, so its marker touches the plot's top edge - see the flip test.
  const twoSeries: ChartData = {
    ...column,
    kind: "line",
    series: [
      column.series[0]!,
      {
        name: "t",
        values: [10, 60, 30],
        labels: ["10", "60", "30"],
        colors: ["#B27E54", "#B27E54", "#B27E54"],
      },
    ],
  };

  it("rebuilds each data point as an editable marker and joins consecutive points", () => {
    const out = layoutChart({ ...column, kind: "line" }, box);
    const markers = out.filter(
      (item) => item.kind === "ellipse" && item.name.startsWith("marker"),
    );
    const lines = out.filter(
      (item) => item.kind === "line" && item.name.startsWith("line"),
    );
    expect(markers).toHaveLength(3);
    expect(lines).toHaveLength(2);
    expect([...markers, ...lines].every((item) => inside(item.box))).toBe(true);
  });

  it("centres each marker on its category label, Excel's own axis centring", () => {
    const out = layoutChart({ ...column, kind: "line" }, box);
    const markers = out.filter(
      (item): item is Ellipse => item.kind === "ellipse",
    );
    const categories = texts(out).filter((t) => t.name.startsWith("category"));
    expect(markers).toHaveLength(categories.length);
    markers.forEach((marker, i) => {
      const markerCentre = marker.box.left + marker.box.width / 2;
      const label = categories[i]!;
      const labelCentre = label.box.left + label.box.width / 2;
      expect(markerCentre).toBeCloseTo(labelCentre, 5);
    });
  });

  it("normalises every connector to a non-negative box and flags a rising segment", () => {
    // column.series[0] is 100, -50, 200: point 0 to 1 falls, point 1 to 2
    // rises. valueY maps a bigger value to a smaller top, so the naive box
    // from a rising pair (previous.top, dy = point.top - previous.top) comes
    // out with a negative height unless it is normalised.
    const out = layoutChart({ ...column, kind: "line" }, box);
    const connectors = lineShapes(out).filter((l) => l.name.startsWith("line"));
    expect(connectors).toHaveLength(2);
    expect(connectors.every((l) => l.box.width >= 0 && l.box.height >= 0)).toBe(
      true,
    );
    expect(connectors.every((l) => inside(l.box))).toBe(true);
    expect(connectors[0]!.rising).toBe(false); // 100 -> -50
    expect(connectors[1]!.rising).toBe(true); // -50 -> 200
  });

  it("labels every point, one label per point per series, inside the box", () => {
    const out = layoutChart(twoSeries, box);
    const markers = out.filter(
      (item): item is Ellipse => item.kind === "ellipse",
    );
    const labels = texts(out).filter((t) => t.name.startsWith("label"));

    expect(markers).toHaveLength(6);
    expect(labels).toHaveLength(6);
    expect(labels.map((t) => t.text).sort()).toEqual(
      ["10", "100", "200", "30", "60", "(50)"].sort(),
    );
    expect(labels.every((t) => inside(t.box))).toBe(true);
  });

  it("flips a label below its marker when the marker sits on the plot's top edge", () => {
    const out = layoutChart({ ...column, kind: "line" }, box);
    const markers = out.filter(
      (item): item is Ellipse => item.kind === "ellipse",
    );
    const labels = texts(out).filter((t) => t.name.startsWith("label"));

    // Point 0.2 is value 200, the scale max: its marker is clamped to the
    // plot's own top edge, so its label must sit below the marker instead of
    // spilling out above the box like a normally placed label would.
    const topMarker = markers.find((m) => m.name === "marker 0.2")!;
    const topLabel = labels.find((t) => t.name === "label 0.2")!;
    expect(topLabel.box.top).toBeCloseTo(
      topMarker.box.top + topMarker.box.height,
      5,
    );

    const midMarker = markers.find((m) => m.name === "marker 0.0")!;
    const midLabel = labels.find((t) => t.name === "label 0.0")!;
    expect(midLabel.box.top + midLabel.box.height).toBeCloseTo(
      midMarker.box.top,
      5,
    );

    expect(inside(topLabel.box)).toBe(true);
    expect(inside(midLabel.box)).toBe(true);
  });
});

describe("layoutChart waterfall", () => {
  const bridge: ChartData = {
    ...column,
    kind: "waterfall",
    categories: ["Open", "Price", "Cost", "Close"],
    series: [
      {
        name: "s",
        values: [100, 20, -30, 90],
        labels: ["100", "20", "(30)", "90"],
        colors: ["#14213D", "#2EC4B6", "#B00020", "#14213D"],
      },
    ],
  };
  it("floats the steps, connects the bars at the running level and labels a fall below its bar", () => {
    const out = layoutChart(bridge, box);
    const bars = rects(out);
    const baseline = out.find((x) => x.name === "baseline")!;
    expect(bars[0]!.box.top + bars[0]!.box.height).toBeCloseTo(
      baseline.box.top,
      0,
    );
    expect(bars[1]!.box.top + bars[1]!.box.height).toBeCloseTo(
      bars[0]!.box.top,
      0,
    );
    expect(bars[2]!.box.top).toBeCloseTo(bars[1]!.box.top, 0);
    expect(out.filter((x) => x.name.startsWith("connector"))).toHaveLength(3);
    const fall = texts(out).find((t) => t.name === "label 0.2")!;
    expect(fall.box.top).toBeGreaterThanOrEqual(
      bars[2]!.box.top + bars[2]!.box.height - 0.01,
    );
  });
});

describe("layoutChart pie", () => {
  const pie: ChartData = {
    ...column,
    kind: "pie",
    categories: ["A", "B", "C", "D"],
    series: [
      {
        name: "s",
        values: [1, 1, 2, 0],
        labels: ["1", "1", "2", "0"],
        colors: ["#14213D", "#2EC4B6", "#B27E54", "#000000"],
      },
    ],
  };
  it("cuts wedges clockwise from 12 o'clock, skips a zero slice and closes the circle", () => {
    const out = layoutChart(pie, box);
    const wedges = out.filter((x): x is Wedge => x.kind === "wedge");
    expect(wedges.map((w) => [w.start, w.end])).toEqual([
      [-90, 0],
      [0, 90],
      [90, -90],
    ]);
    expect(wedges.every((w) => w.box.width === w.box.height)).toBe(true);
    expect(
      texts(out)
        .filter((t) => t.name.startsWith("label"))
        .map((t) => t.text),
    ).toEqual(["1", "1", "2"]);
    expect(
      texts(out)
        .filter((t) => t.name.startsWith("legend"))
        .map((t) => t.text),
    ).toEqual(["A", "B", "C", "D"]);
  });
  it("draws a lone 100 % slice as an ellipse", () => {
    const one: ChartData = {
      ...pie,
      categories: ["A", "B"],
      series: [
        {
          name: "s",
          values: [5, 0],
          labels: ["5", "0"],
          colors: ["#14213D", "#000000"],
        },
      ],
    };
    expect(layoutChart(one, box).some((x) => x.kind === "ellipse")).toBe(true);
  });
});

describe("chartSize", () => {
  it("scales a wide chart down to the content width and never under the minimum", () => {
    expect(chartSize({ width: 1200, height: 600 }, 888)).toEqual({
      width: 888,
      height: 444,
    });
    expect(chartSize({ width: 100, height: 50 }, 888)).toEqual({
      width: 200,
      height: 120,
    });
    expect(chartSize({ width: 420, height: 225 }, 888)).toEqual({
      width: 420,
      height: 225,
    });
  });
});
