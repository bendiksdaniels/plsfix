// The chart data a link can carry: what passes the validator, what the caps
// refuse, and how an angle is brought onto the range PowerPoint reads back.

import { describe, expect, it } from "vitest";
import { isChartData, normalizeAngle, type ChartData } from "./chart-model";

const column: ChartData = {
  v: 1,
  kind: "column",
  title: "Revenue",
  categories: ["2024A", "2025E"],
  series: [
    {
      name: "Revenue",
      values: [100, 120],
      labels: ["100", "120"],
      colors: ["#2EC4B6", "#2EC4B6"],
    },
  ],
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
};

const series = column.series[0]!;

describe("isChartData", () => {
  it("accepts a consistent column chart", () => {
    expect(isChartData(column)).toBe(true);
  });

  it("refuses a series whose labels do not match its values", () => {
    expect(
      isChartData({ ...column, series: [{ ...series, labels: ["100"] }] }),
    ).toBe(false);
  });

  it("refuses more categories than the cap and fewer than two", () => {
    const many = Array.from({ length: 41 }, (_, index) => String(index));
    expect(
      isChartData({
        ...column,
        categories: many,
        series: [
          {
            name: "s",
            values: many.map(Number),
            labels: many,
            colors: many.map(() => "#000000"),
          },
        ],
      }),
    ).toBe(false);
    expect(
      isChartData({
        ...column,
        categories: ["a"],
        series: [
          { name: "s", values: [1], labels: ["1"], colors: ["#000000"] },
        ],
      }),
    ).toBe(false);
  });

  it("refuses a pie with two series or more than twelve slices", () => {
    expect(isChartData({ ...column, kind: "pie" })).toBe(true);
    expect(
      isChartData({ ...column, kind: "pie", series: [series, series] }),
    ).toBe(false);
    const thirteen = Array.from({ length: 13 }, (_, index) => `s${index}`);
    expect(
      isChartData({
        ...column,
        kind: "pie",
        categories: thirteen,
        series: [
          {
            name: "s",
            values: thirteen.map(() => 1),
            labels: thirteen.map(() => "1"),
            colors: thirteen.map(() => "#000000"),
          },
        ],
      }),
    ).toBe(false);
  });

  it("refuses an unknown kind, a bad colour, a non-finite value and a long title", () => {
    expect(isChartData({ ...column, kind: "line" })).toBe(true);
    expect(
      isChartData({
        ...column,
        series: [{ ...series, colors: ["red", "#2EC4B6"] }],
      }),
    ).toBe(false);
    expect(
      isChartData({
        ...column,
        series: [{ ...series, values: [100, Number.NaN] }],
      }),
    ).toBe(false);
    expect(isChartData({ ...column, title: "x".repeat(81) })).toBe(false);
    expect(isChartData({ ...column, title: null })).toBe(true);
  });

  it("allows overlap only as true", () => {
    expect(isChartData({ ...column, kind: "bar", overlap: true })).toBe(true);
    expect(isChartData({ ...column, kind: "bar", overlap: false })).toBe(false);
  });

  it("refuses anything that is not a record", () => {
    expect(isChartData(null)).toBe(false);
    expect(isChartData("chart")).toBe(false);
  });
});

describe("normalizeAngle", () => {
  it("maps onto (-180, 180] the way PowerPoint reads angles back", () => {
    expect(normalizeAngle(200)).toBe(-160);
    expect(normalizeAngle(269.9)).toBeCloseTo(-90.1);
    expect(normalizeAngle(360)).toBe(0);
    expect(normalizeAngle(-90)).toBe(-90);
    expect(normalizeAngle(180)).toBe(180);
    expect(normalizeAngle(-180)).toBe(180);
    expect(normalizeAngle(0)).toBe(0);
  });
});
