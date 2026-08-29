// The unused-style scrubber against the strict fake host: which custom styles
// one pass over the workbook calls unused, which sheets it refuses to read, and
// what the delete does with a workbook it could not read in full.

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

async function boot(sheets: string[] = ["Model", "Data"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
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

// Two custom styles beside Excel's own Normal: one worn by a cell on a hidden
// sheet, one worn by nobody.
function seedStyles(): void {
  helpers.addStyle("Assumption");
  helpers.addStyle("Old header");
  helpers.seed("Data!B2", [["Growth"]]);
  helpers.setStyle("Data!B2", "Assumption");
  helpers.sheet("Data").visibility = "Hidden";
}

// A four-cell block, which is over the cap the scans below run with.
function seedBigSheet(): void {
  helpers.seed("Model!A1", [
    [1, 2],
    [3, 4],
  ]);
}

beforeEach(async () => {
  await boot();
});

describe("list unused styles", () => {
  it("lists the style nobody wears and counts a hidden sheet as use", async () => {
    seedStyles();

    expect(await smt.listUnusedStyles()).toEqual({
      unused: ["Old header"],
      total: 3,
      skippedSheets: [],
    });
  });

  it("never offers a built-in style, however unused", async () => {
    helpers.addStyle("Comma", true);

    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 2,
      skippedSheets: [],
    });
  });

  it("skips a sheet whose used range is past the cap and says which", async () => {
    seedStyles();
    seedBigSheet();

    expect(await smt.listUnusedStyles(3)).toEqual({
      unused: ["Old header"],
      total: 3,
      skippedSheets: ["Model"],
    });
  });

  it("reads an empty workbook without touching a null range", async () => {
    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 1,
      skippedSheets: [],
    });
  });

  // Style properties are the heaviest read of the three workbook scans, and no
  // single sheet has to be large for the request to be: the running total is
  // what stops it, and the sheets left out are named.
  it("stops at the scan cap and names the sheets it did not read", async () => {
    await boot(["Model", "Data", "Notes"]);
    helpers.addStyle("Assumption");
    for (const sheet of ["Model", "Data", "Notes"]) {
      helpers.seed(`${sheet}!A1`, [
        ["a", "b"],
        ["c", "d"],
      ]);
    }
    helpers.setStyle("Notes!A1", "Assumption");

    // Two sheets fit under the total; the third is skipped, so the style only
    // it wears reads as unused and the delete has to refuse.
    expect(await smt.listUnusedStyles(4, 8)).toEqual({
      unused: ["Assumption"],
      total: 2,
      skippedSheets: ["Notes"],
    });
    expect(
      await rejects(() => smt.deleteUnusedStyles(["Assumption"], 4, 8)),
    ).toBe("styles: some sheets were too large to scan");
    expect(workbook.styles).toHaveLength(2);
  });
});

describe("delete unused styles", () => {
  it("deletes the style and leaves the rest of the table alone", async () => {
    seedStyles();

    expect(await smt.deleteUnusedStyles(["Old header"])).toBe(1);
    expect(workbook.styles.map((style) => style.name)).toEqual([
      "Normal",
      "Assumption",
    ]);
  });

  it("leaves a style a cell is still wearing where it is", async () => {
    seedStyles();

    expect(await smt.deleteUnusedStyles(["Assumption", "Old header"])).toBe(1);
    expect(workbook.styles.map((style) => style.name)).toEqual([
      "Normal",
      "Assumption",
    ]);
  });

  it("refuses the whole batch when a sheet was too large to scan", async () => {
    seedStyles();
    seedBigSheet();

    expect(await rejects(() => smt.deleteUnusedStyles(["Old header"], 3))).toBe(
      "styles: some sheets were too large to scan",
    );
    expect(workbook.styles).toHaveLength(3);
  });
});
