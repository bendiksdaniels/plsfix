// A block that only LOOKS empty: a formula showing "" (=IF(...,"")) is the
// modeller's work. requireEmptyBlock guards every block a tool writes beside the
// selection; comps stats stands in for all of them here.

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

const BLANK_FORMULA = '=IF(B2>100,B2,"")';
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
  helpers = installFakeHost({ sheets: ["Model"] }).helpers;
  smt = await import("../../src/excel");
  helpers.seed("Model!A1", [
    ["Company", "EV/EBITDA", "Listed"],
    ["Alpha", 8.1, "yes"],
    ["Beta", 9.4, "yes"],
    ["Gamma", 7.2, "no"],
  ]);
  helpers.select("Model!A1:C4");
});

describe("a block that only looks empty", () => {
  it("refuses to write over a formula that shows nothing", async () => {
    helpers.seed("Model!B8", [[{ formula: BLANK_FORMULA, value: "" }]]);

    expect(await rejects(() => smt.insertCompsStats())).toBe(NOT_EMPTY);
    expect(helpers.formula("Model!B8")).toBe(BLANK_FORMULA);
    expect(helpers.value("Model!A6")).toBe("");
  });

  it("refuses the same way on every press, and leaves the formula alone", async () => {
    helpers.seed("Model!C11", [[{ formula: BLANK_FORMULA, value: "" }]]);

    for (let press = 0; press < 3; press += 1) {
      expect(await rejects(() => smt.insertCompsStats())).toBe(NOT_EMPTY);
    }
    expect(helpers.formula("Model!C11")).toBe(BLANK_FORMULA);
  });

  it("still writes when the six rows hold nothing at all", async () => {
    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 3 rows",
    );
  });
});
