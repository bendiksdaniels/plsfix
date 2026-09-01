import { describe, expect, it } from "vitest";
import type { ChartData } from "./link/chart-model";
import {
  chartSize,
  layoutChart,
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
const inside = (b: {
  left: number;
  top: number;
  width: number;
  height: number;
}) =>
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
