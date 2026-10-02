// The pie leader-lines guard on formatSelectedChart (src/excel/charts.ts,
// about line 202): ChartDataLabels.showLeaderLines is ExcelApi 1.19, but
// ChartSeries.showLeaderLines is 1.9 - the same floor the rest of chart
// formatting already requires - so a host below 1.19 should still keep a
// pie's leader lines through the series-level property.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
});

describe("pie leader lines below ExcelApi 1.19", () => {
  it("sets the series flag on every series but not the label one below 1.19", async () => {
    helpers.setSupported((_set, version) => version !== "1.19");
    const chart = helpers.addChart("Model", {
      chartType: "Pie",
      seriesCount: 2,
    });
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();

    expect(chart.series[0]?.showLeaderLines).toBe(true);
    expect(chart.series[1]?.showLeaderLines).toBe(true);
    expect(chart.dataLabels.showLeaderLines).toBeUndefined();
  });

  it("sets both the series and the label flag on a full 1.19 host", async () => {
    const chart = helpers.addChart("Model", {
      chartType: "Pie",
      seriesCount: 1,
    });
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();

    expect(chart.series[0]?.showLeaderLines).toBe(true);
    expect(chart.dataLabels.showLeaderLines).toBe(true);
  });

  it("sets neither flag for a chart type with no leader lines", async () => {
    const chart = helpers.addChart("Model", {
      chartType: "ColumnClustered",
      seriesCount: 1,
    });
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();

    expect(chart.series[0]?.showLeaderLines).toBeUndefined();
    expect(chart.dataLabels.showLeaderLines).toBeUndefined();
  });
});
