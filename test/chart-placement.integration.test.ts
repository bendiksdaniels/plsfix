// Where an inserted chart lands. The complaint this suite is written for is a
// chart dropped on top of the model, so every case here asserts the corner the
// chart was moved to, against a fake grid of 64pt columns and 15pt rows: a
// 480x288 chart needs 8 columns and 20 rows of air.

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

beforeEach(async () => {
  await boot();
});

// The chart the add-in just inserted: any chart a case seeds is put there
// first, so the newest one is the last in the workbook's list.
function corner(): { left: number; top: number } {
  const chart = workbook.charts[workbook.charts.length - 1];
  if (!chart) throw new Error("no chart");
  return { left: chart.left ?? 0, top: chart.top ?? 0 };
}

// EBITDA bridge at A1, the table at A2:B5: the chart's anchor is 4 rows of two
// columns starting at row 1.
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

describe("a waterfall's corner", () => {
  it("lands right of its table when that block is empty", async () => {
    seedBridge();
    await smt.insertWaterfall();
    // Column D (0 + 2 columns + the gap column), the table's own first row.
    expect(corner()).toEqual({ left: 3 * 64, top: 1 * 15 });
  });

  it("drops below the table when data sits to the right", async () => {
    seedBridge();
    helpers.seed("Model!E3", [[1]]);
    await smt.insertWaterfall();
    expect(corner()).toEqual({ left: 0, top: 6 * 15 });
  });

  it("goes below the used range when both blocks hold data", async () => {
    seedBridge();
    helpers.seed("Model!E3", [[1]]);
    helpers.seed("Model!C10", [[1]]);
    await smt.insertWaterfall();
    // The sheet's values end at row 10, so the block starts one row under it.
    expect(corner()).toEqual({ left: 0, top: 11 * 15 });
  });

  it("never lands on another chart", async () => {
    seedBridge();
    helpers.addChart("Model", {
      name: "Existing",
      left: 3 * 64,
      top: 15,
      width: 200,
      height: 40,
    });
    await smt.insertWaterfall();
    expect(corner()).toEqual({ left: 0, top: 6 * 15 });
  });

  it("counts a merely formatted block as free", async () => {
    seedBridge();
    helpers.setFill("Model!E3", { color: "#FFEECC", pattern: "Solid" });
    await smt.insertWaterfall();
    expect(corner()).toEqual({ left: 3 * 64, top: 1 * 15 });
  });

  it("leaves the chart where Excel dropped it without range geometry", async () => {
    await boot();
    helpers.setSupported((_set, version) => version !== "1.10");
    seedBridge();
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90; Excel placed it",
    );
    expect(corner()).toEqual({ left: 0, top: 0 });
  });
});

describe("a tornado's corner", () => {
  // Title and base at A1, the driver table at A2:C5; the chart's anchor is the
  // helper block the tornado writes at D2:F5.
  function seedDrivers(): void {
    helpers.seed("Model!A1", [["EBITDA sensitivity", 100]]);
    helpers.seed("Model!A2", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
      ["Mix", 95, 105],
    ]);
    helpers.select("Model!A2:C5");
  }

  it("lands right of the helper block when that is empty", async () => {
    seedDrivers();
    await smt.insertTornado();
    expect(corner()).toEqual({ left: 7 * 64, top: 1 * 15 });
  });

  it("drops below the helper block when data sits to the right", async () => {
    seedDrivers();
    helpers.seed("Model!H3", [[1]]);
    await smt.insertTornado();
    expect(corner()).toEqual({ left: 3 * 64, top: 6 * 15 });
  });

  it("goes below the used range when both blocks hold data", async () => {
    seedDrivers();
    helpers.seed("Model!H3", [[1]]);
    helpers.seed("Model!E10", [[1]]);
    await smt.insertTornado();
    expect(corner()).toEqual({ left: 3 * 64, top: 11 * 15 });
  });

  it("never lands on another chart", async () => {
    seedDrivers();
    helpers.addChart("Model", {
      name: "Existing",
      left: 7 * 64,
      top: 15,
      width: 200,
      height: 40,
    });
    await smt.insertTornado();
    expect(corner()).toEqual({ left: 3 * 64, top: 6 * 15 });
  });

  it("says so when the host cannot place it", async () => {
    await boot();
    helpers.setSupported((_set, version) => version !== "1.10");
    seedDrivers();
    expect(await smt.insertTornado()).toBe(
      "Tornado added: 3 drivers, base 100; Excel placed it",
    );
  });
});

describe("the placement budget", () => {
  it("answers in two syncs", async () => {
    seedBridge();
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

      const before = helpers.syncCount();
      expect(await place.placeChartBeside(context, sheet, chart, anchor)).toBe(
        true,
      );
      expect(helpers.syncCount() - before).toBe(2);
    });
  });
});
