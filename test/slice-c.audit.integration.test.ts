// Audit suite for the sheet charts: the hosts that refuse a chartex surface,
// the selection a modeller makes by clicking a column header, a chart anchored
// on rows the host cannot size, and the chart properties each Excel build has.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], ...options });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

function seedBridge(): void {
  helpers.seed("Model!A1", [["EBITDA bridge"]]);
  helpers.seed("Model!A2", [
    ["Opening", 100],
    ["Price", 20],
    ["Cost", -30],
    ["Closing", 90],
  ]);
  helpers.select("Model!A2:B5");
}

function seedDrivers(): void {
  helpers.seed("Model!A1", [["EBITDA sensitivity", 100]]);
  helpers.seed("Model!A2", [
    ["Driver", "Low", "High"],
    ["Volume", 90, 115],
    ["Price", 60, 140],
  ]);
  helpers.select("Model!A2:C4");
}

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------

describe("restyling a chart the host refuses a surface on", () => {
  // Excel for the web answers UnsupportedOperation to chart.format.font and
  // roundedCorners on chartex charts (a waterfall, a treemap, a funnel), the
  // same refusal the insert already tolerates. Restyling one must not lose the
  // branding with it.
  it("brands an existing waterfall instead of failing whole", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedBridge();
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.seriesCount = 1;
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();

    // The surface went down with the host; everything the restyle is for
    // landed anyway.
    expect(chart.font).toEqual({});
    expect(chart.roundedCorners).toBeUndefined();
    expect(chart.borderLineStyle).toBe("None");
    expect(chart.legend.position).toBe("Bottom");
    expect(chart.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
    });
    expect(chart.series[0]?.fillColor).toBeDefined();
  });

  it("still paints the surface where the host takes it", async () => {
    seedBridge();
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.seriesCount = 1;
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();
    expect(chart.roundedCorners).toBe(false);
    expect(chart.font.name).toBeDefined();
  });
});

// ---------------------------------------------------------------------------

describe("a clicked column header", () => {
  // The cap answers before a single value is read: two whole columns are two
  // million cells, and the shape rules only run after the read.
  it("is refused by the waterfall before its values are read", async () => {
    seedBridge();
    helpers.select("Model!A1:B1048576");

    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "Waterfall supports up to 5,000 selected cells at once.",
    );
  });

  it("is refused by the tornado before its values are read", async () => {
    seedDrivers();
    helpers.select("Model!A1:C1048576");

    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado supports up to 5,000 selected cells at once.",
    );
  });

  it("still names the point cap for a table under the cell cap", async () => {
    helpers.select("Model!A1:B101");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "A bridge chart supports up to 100 points.",
    );
  });
});

// ---------------------------------------------------------------------------

describe("a chart anchored on rows the host cannot size", () => {
  // Excel.RangeFormat.rowHeight and columnWidth answer null when the rows or
  // columns of the range are not all the same size, which a label column
  // beside a value column always is
  // (learn.microsoft.com/javascript/api/excel/excel.rangeformat).
  it("places it beside the table rather than under the whole sheet", async () => {
    seedBridge();
    // Outside the eight columns and twenty rows the chart really covers, well
    // inside the block a size of one point computes: a chart sized off a null
    // row height reads this block as occupied and drops under the sheet.
    helpers.seed("Model!M30", [[1]]);
    const place = await import("../src/excel/chart-place");

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Model");
      const anchor = sheet.getRange("A2:B5");
      const chart = sheet.charts.add(
        Excel.ChartType.columnClustered,
        anchor,
        Excel.ChartSeriesBy.auto,
      );
      await context.sync();

      // The same range, with the host answering null for both sizes.
      const mixed = new Proxy(anchor, {
        get(target, property, receiver) {
          if (property === "format") {
            return { rowHeight: null, columnWidth: null };
          }
          const value = Reflect.get(target, property, receiver) as unknown;
          return typeof value === "function" ? value.bind(target) : value;
        },
      });

      expect(await place.placeChartBeside(context, sheet, chart, mixed)).toBe(
        true,
      );
      await context.sync();
    });

    // Column D, the table's own first row: the block right of the anchor.
    const placed = workbook.charts[workbook.charts.length - 1];
    expect({ left: placed?.left, top: placed?.top }).toEqual({
      left: 3 * 64,
      top: 1 * 15,
    });
  });
});

describe("a sheet with no free block at all", () => {
  it("steps the chart under the chart in the way", async () => {
    seedBridge();
    // Right of the table and below it both hold values, so the only candidate
    // left is the block under the used range - and a chart is sitting on it.
    helpers.seed("Model!E3", [[1]]);
    helpers.seed("Model!C10", [[1]]);
    helpers.addChart("Model", {
      name: "Existing",
      left: 0,
      top: 11 * 15,
      width: 200,
      height: 40,
    });

    await smt.insertWaterfall();
    const chart = workbook.charts[workbook.charts.length - 1];
    // Dropped under the chart that blocked the last candidate, gap and all.
    expect(chart?.top).toBe(11 * 15 + 40 + 12);
  });
});

// ---------------------------------------------------------------------------

describe("an Excel build below the chart APIs", () => {
  // Office 2019 runs the shared runtime at ExcelApi 1.8, so every property
  // above it has to be asked for by its own requirement set.
  it("adds the waterfall without its 1.9 connector lines", async () => {
    helpers.setSupported((_set, version) => version !== "1.9");
    seedBridge();

    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90",
    );
    // ChartSeries.showConnectorLines is ExcelApi 1.9
    // (learn.microsoft.com/javascript/api/excel/excel.chartseries).
    expect(workbook.charts[0]?.series[0]?.showConnectorLines).toBeUndefined();
    expect(workbook.charts[0]?.series[0]?.pointColors).toBeDefined();
  });

  it("restyles a pie without its 1.19 leader lines", async () => {
    helpers.setSupported((_set, version) => version !== "1.19");
    seedBridge();
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.chartType = "Pie";
    chart.seriesCount = 1;
    chart.axes = { category: {}, value: {} };
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();
    // ChartDataLabels.showLeaderLines is ExcelApi 1.19; the 1.9 property of
    // that name is ChartSeries.showLeaderLines
    // (learn.microsoft.com/javascript/api/excel/excel.chartdatalabels).
    expect(chart.dataLabels.showLeaderLines).toBeUndefined();
    expect(chart.dataLabels.position).toBe("OutsideEnd");
    expect(chart.legend.visible).toBe(true);
  });

  it("refuses to reconcile without the RangeAreas selection", async () => {
    helpers.setSupported((_set, version) => version !== "1.9");
    helpers.seed("Model!A1", [[10], [20], [30]]);
    helpers.select("Model!A1:A3");

    // worksheet.getRanges and RangeAreas.select are ExcelApi 1.9
    // (learn.microsoft.com/javascript/api/excel/excel.worksheet).
    expect(await rejects(() => smt.reconcileSelection(30, 0))).toBe(
      "Find a combination needs a newer Excel build.",
    );
  });
});
