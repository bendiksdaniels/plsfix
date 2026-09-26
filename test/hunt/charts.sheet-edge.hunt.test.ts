// Attacks: a chart insert whose helper block just barely fits beside the
// selection (requireRoomBeside passes) but whose OWN width leaves no
// candidate at all for placeChartBeside - not "blocked by another chart"
// (already covered: "a sheet with no free block at all" in
// slice-c.audit.integration.test.ts, which still finds a spot via
// dropBelow) but every one of candidateCorners's three slots filtered out
// by SHEET_COLUMNS before readSlots ever runs. The pass-1 lens's sheet-edge
// case (column XFD) for the placement side of the tornado and the football
// field.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

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
// zero. The default chart is 480pt wide over 64pt columns, 8 columns, so
// none of the three placement candidates - right of the block, below it, or
// below the sheet - has room for it from column XFB onward.
describe("a helper block that fills to the sheet's last column", () => {
  it("tornado: says Excel placed it instead of throwing or misplacing", async () => {
    helpers.seed("Model!XEY1", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!XEY1:XFA2");

    const message = await smt.insertTornado();

    expect(message).toBe("Tornado added: 2 drivers, base 100; Excel placed it");
    expect(workbook.charts).toHaveLength(1);
    // Never moved: chart.left/top are only ever written on a successful
    // placement, so an unplaced chart still reads the fake's own default.
    expect(workbook.charts[0]?.left).toBeUndefined();
    // The helper block itself still landed: only the CHART's placement was
    // impossible, not the write beside the selection.
    expect(helpers.value("Model!XFB1")).toBe("Driver");
  });

  it("football field: says Excel placed it instead of throwing or misplacing", async () => {
    helpers.seed("Model!XEY1", [
      ["DCF", 90, 130],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!XEY1:XFA2");

    const message = await smt.insertFootballField();

    expect(message).toBe("Football field added: 2 ranges; Excel placed it");
    expect(workbook.charts).toHaveLength(1);
    expect(workbook.charts[0]?.left).toBeUndefined();
    expect(helpers.value("Model!XFB1")).toBe("Method");
  });
});
