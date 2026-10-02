// Audit suite for the chart payload's validator: what a decoder must refuse.
// The chart data arrives off the relay as whatever bytes decrypted, so every
// field it is trusted for has to be checked here or the slide draws with it.

import { describe, expect, it } from "vitest";
import {
  chartCapIssue,
  isChartData,
  seriesCountIssue,
  type ChartData,
} from "./chart-model";

const valid: ChartData = {
  v: 1,
  kind: "column",
  title: "Revenue",
  categories: ["2024A", "2025E"],
  series: [
    {
      name: "Revenue",
      values: [120, 140],
      labels: ["120", "140"],
      colors: ["#B27E54", "#B27E54"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

function withField(field: string, value: unknown): unknown {
  return { ...valid, [field]: value };
}

describe("isChartData refuses what a slide cannot wear", () => {
  it("takes the shape the reader builds", () => {
    expect(isChartData(valid)).toBe(true);
  });

  it("refuses a brand field that is not a string", () => {
    // The font, the ink and the title colour go straight onto a shape: a
    // number here would be written to PowerPoint as one.
    expect(isChartData(withField("font", 12))).toBe(false);
    expect(isChartData(withField("ink", null))).toBe(false);
    expect(isChartData(withField("titleColor", ["#14213D"]))).toBe(false);
  });

  it("refuses a value that is not a finite number", () => {
    const gap = {
      ...valid,
      series: [{ ...valid.series[0]!, values: [120, Number.NaN] }],
    };
    expect(isChartData(gap)).toBe(false);
    const endless = {
      ...valid,
      series: [
        { ...valid.series[0]!, values: [120, Number.POSITIVE_INFINITY] },
      ],
    };
    expect(isChartData(endless)).toBe(false);
  });

  it("refuses a colour that is not a hex triplet", () => {
    const named = {
      ...valid,
      series: [{ ...valid.series[0]!, colors: ["#B27E54", "red"] }],
    };
    expect(isChartData(named)).toBe(false);
  });

  it("refuses a series whose labels do not cover its points", () => {
    const short = {
      ...valid,
      series: [{ ...valid.series[0]!, labels: ["120"] }],
    };
    expect(isChartData(short)).toBe(false);
  });
});

describe("the sentence a chart outside the caps carries", () => {
  it("counts one point in the singular, the way the modeller sees it", () => {
    expect(chartCapIssue("column", 1, 1)).toBe(
      "1 point; shapes need at least 2",
    );
    expect(chartCapIssue("column", 0, 1)).toBe(
      "0 points; shapes need at least 2",
    );
  });

  it("names a chart with no series at all before any other cap", () => {
    expect(seriesCountIssue(0)).toBe("no series");
    expect(chartCapIssue("pie", 1, 0)).toBe("no series");
  });

  it("refuses a series that is not an object at all", () => {
    expect(isChartData({ ...valid, series: ["Revenue"] })).toBe(false);
    expect(isChartData({ ...valid, series: [] })).toBe(false);
  });
});
