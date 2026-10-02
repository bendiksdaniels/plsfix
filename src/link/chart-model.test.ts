// The chart data a link can carry: what passes the validator, what the caps
// refuse, and how an angle is brought onto the range PowerPoint reads back.

import { describe, expect, it } from "vitest";
import {
  chartCapIssue,
  CHART_MAX_SERIES,
  isChartData,
  normalizeAngle,
  type ChartData,
} from "./chart-model";

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

// Six series is what the palette paints; a seventh has no colour of its own.
describe("series cap", () => {
  it("accepts six series and refuses a seventh", () => {
    const six = Array.from({ length: 6 }, (_, index) => ({
      ...series,
      name: `s${String(index)}`,
    }));
    expect(CHART_MAX_SERIES).toBe(6);
    expect(isChartData({ ...column, series: six })).toBe(true);
    expect(isChartData({ ...column, series: [...six, series] })).toBe(false);
  });
});

// The sentence the panes show after "as a picture": the first cap the chart
// is outside of, counted the way the modeller sees the chart.
describe("chartCapIssue", () => {
  it("is null for a chart inside every cap", () => {
    expect(chartCapIssue("column", 2, 1)).toBeNull();
    expect(chartCapIssue("pie", 12, 1)).toBeNull();
    expect(chartCapIssue("line", 40, 6)).toBeNull();
  });

  it("names the series, point and slice caps", () => {
    expect(chartCapIssue("column", 2, 0)).toBe("no series");
    expect(chartCapIssue("column", 2, 7)).toBe("7 series; shapes draw up to 6");
    expect(chartCapIssue("column", 1, 1)).toBe(
      "1 point; shapes need at least 2",
    );
    expect(chartCapIssue("column", 41, 1)).toBe(
      "41 points; shapes draw up to 40",
    );
    expect(chartCapIssue("pie", 3, 2)).toBe(
      "a pie with 2 series; pie shapes draw one",
    );
    expect(chartCapIssue("pie", 13, 1)).toBe(
      "13 slices; pie shapes draw up to 12",
    );
  });
});
