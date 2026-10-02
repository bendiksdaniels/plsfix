// A block that only looks empty because a dynamic array spills blanks into it
// (rig 27.09): a spill child answers values "" and formulas "", like a truly
// empty cell, and only valueTypes ("String", not "Empty") tells them apart.
// Writing over it breaks the modeller's array, so the guard must refuse it.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

const SPILL = '=IF(SEQUENCE(3)=1,"Peers","")';
const NOT_EMPTY = "Comps stats need six empty rows under the block.";

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"], web: true }).helpers;
  smt = await import("../../src/excel");
  helpers.seed("Model!A1", [
    ["Company", "EV/EBITDA", "Listed"],
    ["Alpha", 8.1, "yes"],
    ["Beta", 9.4, "yes"],
    ["Gamma", 7.2, "no"],
  ]);
  helpers.select("Model!A1:C4");
});

describe("a block a spill runs through", () => {
  it("refuses comps stats over a spill's blank cells and leaves them alone", async () => {
    // Anchored in the gap row, its two blank children inside the target.
    helpers.spill("Model!B5", SPILL, [["Peers"], [""], [""]]);

    expect(await rejects(() => smt.insertCompsStats())).toBe(NOT_EMPTY);
    expect(helpers.value("Model!B6")).toBe("");
    expect(helpers.formula("Model!B6")).toBe("");
    expect(helpers.value("Model!A6")).toBe("");
  });

  it("refuses a tornado block over a spill's blank cells", async () => {
    helpers.seed("Model!A10", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
    ]);
    helpers.spill("Model!D9", SPILL, [["Peers"], [""], [""], [""]]);
    helpers.select("Model!A10:C12");

    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado: cells to the right of the selection are not empty",
    );
    expect(helpers.value("Model!D10")).toBe("");
  });
});

describe("a block that is empty", () => {
  it("still takes comps stats when nothing stands there", async () => {
    expect(await smt.insertCompsStats()).toContain("Comps stats");
    expect(helpers.formula("Model!B6")).toMatch(/^=MIN\(/);
  });

  it("still takes comps stats over cells that only wear formatting", async () => {
    helpers.setFill("Model!A6:C11", { color: "#FFFF00", pattern: "Solid" });

    expect(await smt.insertCompsStats()).toContain("Comps stats");
    expect(helpers.formula("Model!B6")).toMatch(/^=MIN\(/);
  });
});
