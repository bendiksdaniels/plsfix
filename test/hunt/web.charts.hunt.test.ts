// Chart labels on Excel for the web (rig 27.09): a bar chart refuses a label's
// percentage and bubble size (InvalidOperation), a chartex chart both plus its
// legend key (UnsupportedOperation), each ending its batch. The tornado and
// Chart format set a part only where the chart has it, and still land.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeChart,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import { DEFAULT_SETTINGS } from "../../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], web: true });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../../src/excel");
});

function restyle(chartType: string, seriesCount: number): FakeChart {
  const chart = helpers.addChart("Model", { chartType, seriesCount });
  helpers.setActiveChart(chart);
  return chart;
}

describe("the tornado on the web", () => {
  it("places the chart beside the data and answers its own sentence", async () => {
    helpers.seed("Model!A1", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
    ]);
    helpers.select("Model!A1:C3");

    expect(await smt.insertTornado()).toBe(
      "Tornado added: 2 drivers, base 101.3",
    );

    const chart = workbook.charts[0];
    expect(chart).toMatchObject({ left: 7 * 64, top: 0 });
    expect(chart?.dataLabels).toMatchObject({
      showValue: true,
      showLegendKey: false,
      position: "OutsideEnd",
    });
    expect(chart?.dataLabels.font.name).toBe(DEFAULT_SETTINGS.font);
    expect(chart?.axes.category.reversePlotOrder).toBe(true);
    expect(chart?.series[1]?.fillColor).toBe(DEFAULT_SETTINGS.accent);
  });
});

describe("Chart format on the web", () => {
  it("restyles a clustered bar chart, surface included", async () => {
    const chart = restyle("BarClustered", 2);

    await smt.formatSelectedChart();

    expect(chart.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showLegendKey: false,
      position: "OutsideEnd",
    });
    expect(chart.series.map((series) => series.fillColor)).toHaveLength(2);
    // A bar chart takes the surface on the web.
    expect(chart.font.name).toBe(DEFAULT_SETTINGS.font);
    expect(chart.roundedCorners).toBe(false);
  });

  it("switches a pie's percentage off and a bubble chart's bubble size", async () => {
    const pie = restyle("Pie", 1);
    await smt.formatSelectedChart();
    expect(pie.dataLabels.showPercentage).toBe(false);

    const bubble = restyle("Bubble", 1);
    await smt.formatSelectedChart();
    expect(bubble.dataLabels.showBubbleSize).toBe(false);
  });

  it("restyles a waterfall, leaving its refused surface to the host", async () => {
    const chart = restyle("Waterfall", 1);

    await smt.formatSelectedChart();

    expect(chart.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
    });
    expect(chart.legend.position).toBe("Bottom");
    expect(chart.font).toEqual({});
  });
});
