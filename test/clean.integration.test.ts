// Clean past the data: the used range Excel remembers against the one the
// values actually fill. Rows below and columns right of the values are the
// surplus a saved workbook carries around; they go, unless the sheet holds a
// chart or a shape, in which case only their formatting does.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["P&L", "Data"], ...options });
  helpers = host.helpers;
  smt = await import("../src/excel");
}

// A host that answers no to one API set and yes to everything else.
function without(apiSet: string): FakeHostOptions {
  return {
    isSetSupported: (set, version) =>
      !(set === "ExcelApi" && version === apiSet),
  };
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// A small model in A1:C3, with formatting Excel still counts as "used" far
// below it and a couple of columns to its right.
function seedStrays(rows: number, columns: number): void {
  helpers.seed("P&L!A1", [
    ["Revenue", 100, 120],
    ["Cost", -40, -50],
    ["Profit", 60, 70],
  ]);
  if (rows > 0) {
    helpers.setFill(`P&L!A4:C${String(3 + rows)}`, {
      color: "#FFEECC",
      pattern: "Solid",
    });
  }
  if (columns > 0) {
    const last = String.fromCharCode("C".charCodeAt(0) + columns);
    helpers.setFill(`P&L!D1:${last}3`, {
      color: "#FFEECC",
      pattern: "Solid",
    });
  }
}

beforeEach(async () => {
  await boot();
});

describe("clean past the data", () => {
  it("deletes a thousand stray formatted rows and counts them", async () => {
    seedStrays(1000, 0);

    expect(await smt.cleanPastData()).toBe(
      "Removed 1,000 rows past the data on P&L",
    );
    // The data is untouched and the used range is the data again.
    expect(helpers.value("P&L!C3")).toBe(70);
    expect(helpers.fill("P&L!A4").pattern).toBe("None");
  });

  it("counts rows and columns together", async () => {
    seedStrays(1204, 12);

    expect(await smt.cleanPastData()).toBe(
      "Removed 1,204 rows and 12 columns past the data on P&L",
    );
    expect(helpers.fill("P&L!D1").pattern).toBe("None");
  });

  it("counts one row and one column in the singular", async () => {
    seedStrays(1, 1);

    expect(await smt.cleanPastData()).toBe(
      "Removed 1 row and 1 column past the data on P&L",
    );
  });

  it("leaves nothing behind: a second pass finds no surplus", async () => {
    seedStrays(40, 3);
    await smt.cleanPastData();

    expect(await smt.cleanPastData()).toBe("Nothing past the data on P&L.");
  });

  it("keeps the rows and columns when the sheet holds a chart", async () => {
    seedStrays(1000, 2);
    helpers.addChart("P&L");

    expect(await smt.cleanPastData()).toBe(
      "Cleared the formats past the data on P&L; rows and columns kept because the sheet has charts or shapes",
    );
    // The formatting is gone, the rows themselves are still there.
    expect(helpers.fill("P&L!A4").pattern).toBe("None");
    expect(helpers.fill("P&L!D1").pattern).toBe("None");
    expect(helpers.value("P&L!C3")).toBe(70);
  });

  it("says there is nothing to clean on an empty sheet", async () => {
    expect(await smt.cleanPastData()).toBe("Nothing on this sheet to clean.");
  });

  it("says so when the used range is the data range", async () => {
    seedStrays(0, 0);
    expect(await smt.cleanPastData()).toBe("Nothing past the data on P&L.");
  });

  it("leaves a sheet that is all formatting and no values alone", async () => {
    helpers.setFill("P&L!A1:C50", { color: "#FFEECC", pattern: "Solid" });

    expect(await smt.cleanPastData()).toBe("Nothing on this sheet to clean.");
    expect(helpers.fill("P&L!A1").pattern).toBe("Solid");
  });

  it("answers a protected sheet with the pane's own sentence", async () => {
    seedStrays(50, 0);
    helpers.protectSheet("P&L");
    const before = helpers.cellMap("P&L");

    expect(await rejects(() => smt.cleanPastData())).toBe(
      "Clean past the data: this sheet is protected, nothing was changed",
    );
    expect(helpers.cellMap("P&L")).toEqual(before);
  });

  it("cleans the sheet the workbook is looking at, and no other", async () => {
    seedStrays(10, 0);
    helpers.setFill("Data!A1:B40", { color: "#FFEECC", pattern: "Solid" });

    await smt.cleanPastData();
    expect(helpers.fill("Data!A40").pattern).toBe("Solid");
  });

  it("spends no pls,fix undo slot: this is sheet state, not cell state", async () => {
    seedStrays(10, 0);
    await smt.cleanPastData();
    expect(smt.undoTarget()).toBeNull();
  });

  // A merged block is the one thing that tells the two branches apart in the
  // fake: a format clear leaves it standing, a whole-row delete takes it with
  // the rows, which is what Excel does to a merge the deleted rows cover.
  it("deletes through a merged block sitting past the data", async () => {
    seedStrays(10, 0);
    helpers.merge("P&L!A6:B6");

    expect(await smt.cleanPastData()).toBe(
      "Removed 10 rows past the data on P&L",
    );
    expect(helpers.sheet("P&L").merges).toEqual([]);
    expect(helpers.value("P&L!C3")).toBe(70);
    expect(helpers.fill("P&L!A4").pattern).toBe("None");
  });
});

// An un-undoable delete may never run on a guess: a host that cannot be asked
// how many charts or shapes sit on the sheet takes the safe branch, because a
// logo or a text box past the data would be destroyed with the rows.
describe("clean past the data on a host that cannot count drawings", () => {
  const CANNOT_COUNT =
    "Cleared the formats past the data on P&L; rows and columns kept because this Excel cannot count the sheet's charts or shapes";

  it("keeps the rows when shapes cannot be counted (below ExcelApi 1.9)", async () => {
    await boot(without("1.9"));
    seedStrays(1000, 2);
    // The witness for "the rows are still there": a delete would take this
    // merge with them, a format clear leaves it standing.
    helpers.merge("P&L!A6:B6");

    expect(await smt.cleanPastData()).toBe(CANNOT_COUNT);
    expect(helpers.sheet("P&L").merges).toHaveLength(1);
    // The formatting past the data went all the same.
    expect(helpers.fill("P&L!A4").pattern).toBe("None");
    expect(helpers.fill("P&L!D1").pattern).toBe("None");
  });

  it("keeps the rows when charts cannot be counted (below ExcelApi 1.4)", async () => {
    await boot(without("1.4"));
    seedStrays(40, 0);
    helpers.merge("P&L!A6:B6");

    expect(await smt.cleanPastData()).toBe(CANNOT_COUNT);
    expect(helpers.sheet("P&L").merges).toHaveLength(1);
  });

  it("still deletes on a host that can count both and has neither", async () => {
    await boot();
    seedStrays(40, 0);

    expect(await smt.cleanPastData()).toBe(
      "Removed 40 rows past the data on P&L",
    );
  });
});
