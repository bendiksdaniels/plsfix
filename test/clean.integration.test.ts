// Clean past the data: the used range Excel remembers against the one the
// values actually fill. Rows below and columns right of the values are the
// surplus a saved workbook carries around; they go, unless the sheet holds a
// chart or a shape, in which case only their formatting does.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["P&L", "Data"] });
  helpers = host.helpers;
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
});
