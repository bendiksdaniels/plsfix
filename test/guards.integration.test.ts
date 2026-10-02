// The guards that stand between a pane action and the modeller's own cells:
// the broken-name scrubber leaves link anchors alone, the tornado refuses to
// write over a helper block that is not empty, and every flow that needs one
// rectangle says so rather than letting office.js throw on a ctrl-clicked
// selection. Driven end to end against the strict fake host.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
});

// A link whose rows were deleted is a #REF! hidden name by design; the Links
// tab reports it as "Source missing" and owns its removal.
describe("broken names and link anchors", () => {
  beforeEach(() => {
    helpers.addName("Costs", "=Model!#REF!");
    helpers.addName("PLSFIX_LINK_5f3a91c2", "=Model!#REF!");
  });

  it("never lists a link anchor among the broken names", async () => {
    expect(await smt.listBrokenNames()).toEqual(["Costs"]);
  });

  it("never deletes a link anchor with them", async () => {
    expect(await smt.deleteBrokenNames()).toBe(1);
    expect(workbook.names.map((entry) => entry.name)).toEqual([
      "PLSFIX_LINK_5f3a91c2",
    ]);
  });
});

describe("tornado over an occupied helper block", () => {
  beforeEach(() => {
    helpers.seed("Model!A2", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
    ]);
    helpers.select("Model!A2:C4");
  });

  it("refuses when a cell to the right holds anything", async () => {
    helpers.seed("Model!E3", [[100]]);
    await expect(smt.insertTornado()).rejects.toThrow(
      "tornado: cells to the right of the selection are not empty",
    );
    expect(helpers.value("Model!D2")).toBe("");
    expect(workbook.charts).toEqual([]);
  });

  it("writes when the block is free", async () => {
    await smt.insertTornado();
    expect(helpers.value("Model!D2")).toBe("Driver");
  });
});

// office.js: getSelectedRange "will throw an error" on a multi-area selection,
// and it throws with no stage of its own.
describe("a multi-area selection", () => {
  beforeEach(() => {
    helpers.seed("Model!A2", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
    ]);
  });

  it("is refused by the tornado", async () => {
    helpers.selectAreas(["Model!A2:C4", "Model!A6:C8"]);
    await expect(smt.insertTornado()).rejects.toThrow(
      "tornado: select a single range",
    );
  });

  it("is refused by unpivot", async () => {
    helpers.selectAreas(["Model!A2:C4", "Model!A6:C8"]);
    await expect(smt.unpivotSelection()).rejects.toThrow(
      "unpivot: select a single range",
    );
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
      "Model",
      "Data",
    ]);
  });

  it("is refused by the size cycles", async () => {
    helpers.selectAreas(["Model!A2:C4", "Model!A6:C8"]);
    await expect(smt.applyRowHeightCycle()).rejects.toThrow(
      "Row height: select a single range",
    );
    await expect(smt.applyColumnWidthCycle()).rejects.toThrow(
      "Column width: select a single range",
    );
  });

  it("still works as one block", async () => {
    helpers.selectAreas(["Model!A2:C4"]);
    expect(await smt.unpivotSelection()).toBe("Unpivot: 4 rows on Unpivot");
  });
});
