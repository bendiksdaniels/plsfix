// Attacks: a chart insert whose helper block sits at the sheet's right edge
// (column XFD), where the "below" placement corners must pull back to fit
// the grid instead of leaving the chart on the data; and placeChartBeside's
// own empty-sheet fallback, which every other placement test leaves untouched.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import { overlaps } from "../../src/layout";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../../src/excel");
});

// Columns XEY, XEZ, XFA (0-indexed 16378-16380): requireRoomBeside allows the
// three-column helper block right after them (XFB:XFD, 0-indexed
// 16381-16383, the sheet's last three columns) with room to spare of exactly
// zero. The default chart is 480pt wide over 64pt columns, 8 columns, so the
// right-of-anchor corner still has no room from column XFB onward - but the
// below-anchor corner now pulls its column back to 16376 (16384 - 8) to fit,
// landing at row 4 (below the 3-row block), left 16376*64, top 4*15.
const SELECTION_BOX = {
  left: 16378 * 64,
  top: 0,
  width: 3 * 64,
  height: 2 * 15,
};

describe("a helper block that fills to the sheet's last column", () => {
  it("tornado: pulls the placement back to fit, clear of the selection", async () => {
    helpers.seed("Model!XEY1", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!XEY1:XFA2");

    const message = await smt.insertTornado();

    expect(message).toBe("Tornado added: 2 drivers, base 100");
    expect(workbook.charts).toHaveLength(1);
    const chart = workbook.charts[0]!;
    expect(chart).toMatchObject({ left: 16376 * 64, top: 4 * 15 });
    expect(
      overlaps(
        {
          left: chart.left ?? 0,
          top: chart.top ?? 0,
          width: chart.width,
          height: chart.height,
        },
        SELECTION_BOX,
      ),
    ).toBe(false);
    // The helper block itself still landed right of the selection.
    expect(helpers.value("Model!XFB1")).toBe("Driver");
  });

  it("football field: pulls the placement back to fit, clear of the selection", async () => {
    helpers.seed("Model!XEY1", [
      ["DCF", 90, 130],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!XEY1:XFA2");

    const message = await smt.insertFootballField();

    expect(message).toBe("Football field added: 2 ranges");
    expect(workbook.charts).toHaveLength(1);
    const chart = workbook.charts[0]!;
    expect(chart).toMatchObject({ left: 16376 * 64, top: 4 * 15 });
    expect(
      overlaps(
        {
          left: chart.left ?? 0,
          top: chart.top ?? 0,
          width: chart.width,
          height: chart.height,
        },
        SELECTION_BOX,
      ),
    ).toBe(false);
    expect(helpers.value("Model!XFB1")).toBe("Method");
  });
});

// readPlan's own bottom falls back to the anchor's own rowIndex + rowCount
// when the sheet's used range is a null object; every other placement test
// seeds a table first, so the sheet always has something and this fallback
// never runs. A genuinely blank sheet is the one way to reach it.
describe("placeChartBeside on a genuinely empty sheet", () => {
  it("still lands right of the anchor, not over nothing that is not there", async () => {
    const place = await import("../../src/excel/chart-place");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Model");
      const anchor = sheet.getRange("A1:B4");
      const chart = sheet.charts.add(
        Excel.ChartType.columnClustered,
        anchor,
        Excel.ChartSeriesBy.auto,
      );
      await context.sync();

      expect(await place.placeChartBeside(context, sheet, chart, anchor)).toBe(
        true,
      );
    });

    // Right of the anchor (2 columns + the gap column), same row: the sheet
    // has nothing else on it at all, so the very first candidate is free.
    expect(workbook.charts[0]).toMatchObject({ left: 3 * 64, top: 0 });
  });
});
