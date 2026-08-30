// The tornado on whatever the modeller actually selected: a driver table with a
// header row or without one, at the top-left corner of a sheet or halfway down
// it, and on a host too old to shape the bars.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
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

beforeEach(async () => {
  await boot();
});

describe("a driver block anywhere on the sheet", () => {
  it("reads a headed table at the very corner", async () => {
    helpers.seed("Model!A1", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
    ]);
    helpers.select("Model!A1:C3");

    // No row above the corner to hold a base, so the mean of the outcomes is it.
    expect(await smt.insertTornado()).toBe(
      "Tornado added: 2 drivers, base 101.3",
    );
    expect(helpers.value("Model!D1")).toBe("Driver");
    expect(workbook.charts[0]?.title).toBe("Sensitivity");
    // Just the values, at the end of each bar.
    expect(workbook.charts[0]?.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
      showPercentage: false,
      showLegendKey: false,
      position: "OutsideEnd",
    });
  });

  it("labels the bars in the number format the outcomes wear", async () => {
    helpers.seed("Model!A1", [
      ["Driver", "Low", "High"],
      ["Volume", 0.09, 0.115],
      ["Price", 0.06, 0.14],
    ]);
    helpers.setNumberFormat("Model!B2:C3", "0.0%");
    helpers.select("Model!A1:C3");

    await smt.insertTornado();

    // The deltas share the outcomes' unit, so the helper block wears their
    // format and the labels the chart reads off it follow; the header stays.
    expect(helpers.numberFormat("Model!E2")).toBe("0.0%");
    expect(helpers.numberFormat("Model!F3")).toBe("0.0%");
    expect(helpers.numberFormat("Model!E1")).toBe("General");
    expect(helpers.numberFormat("Model!D2")).toBe("General");
  });

  it("reads a header-free table halfway down the sheet", async () => {
    helpers.seed("Model!H40", [["Downside case", 100]]);
    helpers.seed("Model!H41", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!H41:J42");

    expect(await smt.insertTornado()).toBe(
      "Tornado added: 2 drivers, base 100",
    );
    expect(
      ["K41", "L41", "M41"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Driver", "Low", "High"]);
    expect(
      ["K42", "L42", "M42"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Volume", -20, 20]);
    expect(workbook.charts[0]).toMatchObject({
      sourceAddress: "Model!K41:M43",
      title: "Downside case",
    });
  });

  it("refuses a block with no room to its right", async () => {
    // XEZ is the fourth column from the right edge: the helper block would run
    // past XFD.
    helpers.seed("Model!XEZ1", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!XEZ1:XFB2");

    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado: no room to the right of the selection",
    );
  });
});

describe("the tornado on an older host", () => {
  it("says the bars are plain when the shaping API is missing", async () => {
    helpers.setSupported((_set, version) => version !== "1.8");
    helpers.seed("Model!A2", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!A2:C3");

    expect(await smt.insertTornado()).toBe(
      "Tornado added: 2 drivers, base 100; plain bars on this build",
    );
    expect(workbook.charts[0]?.series[0]?.overlap).toBeUndefined();
  });
});
