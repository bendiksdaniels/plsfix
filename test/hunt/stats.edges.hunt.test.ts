// Attacks: what insertCompsStats's plain MIN/PERCENTILE.INC/MEDIAN/AVERAGE/MAX
// formulas actually do with a mixed column, an error value, and a filtered
// table - three cases the brief asks to be answered, not assumed - plus three
// presses in a row.

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

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"] }).helpers;
  smt = await import("../../src/excel");
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

// ---------------------------------------------------------------------------
// A column mixing numbers and text: still counted as numeric (one number is
// enough), and the written range spans every row, text ones included. This is
// sound BECAUSE Excel's MIN/MEDIAN/AVERAGE/MAX/PERCENTILE.INC ignore text and
// blanks inside a range reference - they only choke on text given directly as
// an argument, never on text sitting inside the cells a range points at.
// ---------------------------------------------------------------------------

describe("a column mixing numbers and text", () => {
  it("is read as numeric and the block's range covers the text row too", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
      ["Beta", "n/a"],
      ["Gamma", 9.4],
    ]);
    helpers.select("Model!A1:B4");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 3 rows",
    );
    // The range spans B2:B4 whether or not B3 itself holds a number: Excel's
    // MIN/MEDIAN/AVERAGE/MAX read past a text cell inside a range silently.
    expect(helpers.formula("Model!B6")).toBe("=MIN(B2:B4)");
    expect(helpers.formula("Model!B11")).toBe("=MAX(B2:B4)");
  });
});

// ---------------------------------------------------------------------------
// An error value inside the numbers: the column is still numeric (one real
// number is enough) and the range still covers the error cell - which means
// every one of the six statistics for that column will themselves show the
// error, because MIN/PERCENTILE.INC/MEDIAN/AVERAGE/MAX (unlike their handling
// of text) propagate any #N/A, #REF! etc. found anywhere in a range argument.
// One bad lookup blanks out the whole column's stats; AGGREGATE(fn, 6, range)
// (option 6 = ignore errors) is the honest formula if that is not wanted, and
// this tool does not offer it.
// ---------------------------------------------------------------------------

describe("an error value inside the numeric column", () => {
  it("still spans the row an #N/A sits in, so every statistic would show it", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
      ["Beta", "#N/A"],
      ["Gamma", 9.4],
    ]);
    helpers.select("Model!A1:B4");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 3 rows",
    );
    // Every one of the six formulas spans B2:B4: none of them route around
    // the error cell, so all six would show #N/A once Excel recalculates.
    const formulas = ["B6", "B7", "B8", "B9", "B10", "B11"].map((at) =>
      helpers.formula(`Model!${at}`),
    );
    for (const formula of formulas) {
      expect(String(formula)).toContain("B2:B4");
    }
    expect(helpers.value("Model!B3")).toBe("#N/A");
  });
});

// ---------------------------------------------------------------------------
// A filtered table: hidden rows are read exactly like visible ones, both for
// detecting the numeric columns and for the range the six formulas span -
// range.values does not skip a hidden row, and plain MIN/MEDIAN/etc. (unlike
// SUBTOTAL/AGGREGATE with the "ignore hidden rows" option) never look at
// Range.rowHidden either. So hidden rows ARE included in the answer; a comps
// page that means to reflect only what is showing would want
// AGGREGATE(fn, 5 or 7, range) instead.
// ---------------------------------------------------------------------------

describe("a filtered table", () => {
  it("includes a hidden row's value in both detection and the written range", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
      ["Beta", 50], // filtered out by the modeller, but still a real cell
      ["Gamma", 9.4],
    ]);
    helpers.applyFilter("Model!A1:B4", "Model!3:3");
    helpers.select("Model!A1:B4");

    expect(helpers.value("Model!B3")).toBe(50);
    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 3 rows",
    );
    // The range is B2:B4 - the hidden 50 sits inside it, not excluded.
    expect(helpers.formula("Model!B11")).toBe("=MAX(B2:B4)");
    expect(helpers.formula("Model!B6")).toBe("=MIN(B2:B4)");
  });
});

// ---------------------------------------------------------------------------
// Pressed three times in a row.
// ---------------------------------------------------------------------------

describe("pressed three times in a row", () => {
  it("writes once, then refuses twice more over its own block, block untouched both times", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
      ["Beta", 9.4],
    ]);
    helpers.select("Model!A1:B3");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 2 rows",
    );
    const afterFirst = helpers.cellMap("Model");
    const refusal = "Comps stats need six empty rows under the block.";

    helpers.select("Model!A1:B3");
    expect(await rejects(() => smt.insertCompsStats())).toBe(refusal);
    expect(helpers.cellMap("Model")).toEqual(afterFirst);

    helpers.select("Model!A1:B3");
    expect(await rejects(() => smt.insertCompsStats())).toBe(refusal);
    expect(helpers.cellMap("Model")).toEqual(afterFirst);
  });
});
