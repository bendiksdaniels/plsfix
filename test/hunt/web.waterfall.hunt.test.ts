// The waterfall on Excel for the web (rig 27.09): a chartex chart refuses the
// chart-area font and corners (UnsupportedOperation), and the web ends a batch
// at the first refused statement, so a surface written ahead of the labels
// took the labels down with it. The refused surface must cost nothing else.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
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
  helpers.seed("Model!A1", [["EBITDA bridge"]]);
  helpers.seed("Model!A2", [
    ["Opening", 100],
    ["Price", 20],
    ["Cost", -30],
    ["Closing", 90],
  ]);
  helpers.select("Model!A2:B5");
});

describe("the waterfall on the web", () => {
  it("lands its labels and bar colours while the host refuses the surface", async () => {
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90",
    );

    const chart = workbook.charts[0];
    expect(chart?.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
    });
    expect(chart?.dataLabels.font).toMatchObject({
      name: DEFAULT_SETTINGS.font,
      size: 9,
    });
    expect(chart?.series[0]?.pointColors).toEqual({
      0: DEFAULT_SETTINGS.primary,
      1: DEFAULT_SETTINGS.accent,
      2: DEFAULT_SETTINGS.external,
      3: DEFAULT_SETTINGS.primary,
    });
    // The surface itself: refused on a chartex chart, so never written.
    expect(chart?.font).toEqual({});
    expect(chart?.roundedCorners).toBeUndefined();
  });

  it("still titles, places and connects the bridge", async () => {
    await smt.insertWaterfall();

    const chart = workbook.charts[0];
    expect(chart).toMatchObject({
      chartType: "Waterfall",
      title: "EBITDA bridge",
    });
    expect(chart?.legend.visible).toBe(false);
    expect(chart?.series[0]?.showConnectorLines).toBe(true);
    expect(chart?.left).toBeGreaterThan(0);
  });
});
