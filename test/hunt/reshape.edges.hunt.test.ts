// Attacks: unpivotSelection against the workbook's own sheet-list protection
// (never checked before), plus the header/body edges the shared brief calls
// out - duplicate headers and row keys, a one-row grid, a 200-column-wide
// grid, an error value in the body, a merged row-key column, a 31-character
// source sheet name, and the same press twice.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import { unpivot } from "../../src/reshape";
import { cellAddress } from "../../src/find";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(sheets: string[] = ["Model"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets }).helpers;
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
// The bug: adding the long-table sheet was never checked against workbook
// structure protection, so it either wrote past the lock (fake) or would have
// surfaced Excel's own bare AccessDenied instead of a pls,fix sentence (real
// Excel, which refuses worksheets.add() the same way it refuses a hide).
// ---------------------------------------------------------------------------

describe("workbook structure protection", () => {
  function seedCrossTab(): void {
    helpers.seed("Model!A1", [
      ["", "North", "South"],
      ["Q1", 10, 20],
    ]);
    helpers.select("Model!A1:C2");
  }

  it("names the sheet lock instead of adding a sheet the lock should have refused", async () => {
    seedCrossTab();
    helpers.protectWorkbook();

    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: this workbook's structure is protected, nothing was changed",
    );
    // No phantom "Unpivot" sheet, and the source is exactly as seeded.
    expect(helpers.sheet("Model").name).toBe("Model");
    expect(() => helpers.sheet("Unpivot")).toThrow();
    expect(helpers.value("Model!B2")).toBe(10);
  });
});

// ---------------------------------------------------------------------------
// Duplicate headers and duplicate row keys: a positional unpivot, not a
// group-by, so each one gets its own line rather than being merged or
// silently dropped by a naive "last write wins" map.
// ---------------------------------------------------------------------------

describe("duplicates are not deduplicated", () => {
  it("keeps two columns that share the same header as two separate lines per row", () => {
    expect(
      unpivot([
        ["", "North", "North"],
        ["Q1", 10, 15],
        ["Q2", 30, 35],
      ]),
    ).toEqual([
      ["Q1", "North", 10],
      ["Q1", "North", 15],
      ["Q2", "North", 30],
      ["Q2", "North", 35],
    ]);
  });

  it("keeps two rows that share the same key as two separate lines per column", () => {
    expect(
      unpivot([
        ["", "North", "South"],
        ["Q1", 10, 20],
        ["Q1", 30, 40],
      ]),
    ).toEqual([
      ["Q1", "North", 10],
      ["Q1", "South", 20],
      ["Q1", "North", 30],
      ["Q1", "South", 40],
    ]);
  });

  it("proves the same thing end to end through the adapter", async () => {
    helpers.seed("Model!A1", [
      ["", "North", "North"],
      ["Q1", 10, 15],
    ]);
    helpers.select("Model!A1:C2");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 2 rows on Unpivot");
    expect(helpers.value("Unpivot!C2")).toBe(10);
    expect(helpers.value("Unpivot!C3")).toBe(15);
  });
});

// ---------------------------------------------------------------------------
// A grid with only one row: no data row at all, refused the same way a
// one-column selection is.
// ---------------------------------------------------------------------------

describe("a grid with only one row", () => {
  it("refuses a header with nothing under it", async () => {
    helpers.seed("Model!A1", [["", "North", "South"]]);
    helpers.select("Model!A1:C1");

    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: need a header row, a key column and one column of values",
    );
  });
});

// ---------------------------------------------------------------------------
// A wide grid: 200 columns, one data row - proves every column is read, not
// only the first 26 (up to Z) or some other letter-boundary artefact.
// ---------------------------------------------------------------------------

describe("a wide grid", () => {
  it("unpivots all 199 data columns of a 200-column header, first and last included", async () => {
    const width = 200;
    const header = [
      "",
      ...Array.from({ length: width - 1 }, (_unused, i) => `Col${i + 1}`),
    ];
    const dataRow = [
      "Row1",
      ...Array.from({ length: width - 1 }, (_unused, i) => i + 1),
    ];
    helpers.seed("Model!A1", [header, dataRow]);
    helpers.select(`Model!A1:${cellAddress(1, width - 1)}`);

    expect(await smt.unpivotSelection()).toBe("Unpivot: 199 rows on Unpivot");

    // First data column.
    expect(helpers.value(`Unpivot!${cellAddress(1, 0)}`)).toBe("Row1");
    expect(helpers.value(`Unpivot!${cellAddress(1, 1)}`)).toBe("Col1");
    expect(helpers.value(`Unpivot!${cellAddress(1, 2)}`)).toBe(1);
    // Last data column: nothing truncated it early.
    const lastLine = width - 1;
    expect(helpers.value(`Unpivot!${cellAddress(lastLine, 1)}`)).toBe(
      `Col${lastLine}`,
    );
    expect(helpers.value(`Unpivot!${cellAddress(lastLine, 2)}`)).toBe(lastLine);
  });
});

// ---------------------------------------------------------------------------
// An error value in the body: not blank, so it is not skipped - it travels
// into the long table as the literal error text, the same as any other
// string value would.
// ---------------------------------------------------------------------------

describe("an error value in the body", () => {
  it("carries the error text into the long table instead of dropping the line", async () => {
    helpers.seed("Model!A1", [
      ["", "EV/EBITDA"],
      ["Alpha", 8.1],
      ["Beta", "#N/A"],
    ]);
    helpers.select("Model!A1:B3");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 2 rows on Unpivot");
    expect(helpers.value("Unpivot!C2")).toBe(8.1);
    expect(helpers.value("Unpivot!A3")).toBe("Beta");
    expect(helpers.value("Unpivot!C3")).toBe("#N/A");
  });
});

// ---------------------------------------------------------------------------
// A merged row-key column: only the merge's top-left cell carries the key,
// exactly like a merged header does, so the row under it is skipped rather
// than emitted under a blank key - by the same design as the header case, now
// pinned for the row side too.
// ---------------------------------------------------------------------------

describe("a merged row-key column", () => {
  it("drops the row a vertical merge leaves without its own key", async () => {
    helpers.seed("Model!A1", [
      ["", "Value"],
      ["Group1", 10],
      ["", 20],
    ]);
    helpers.merge("Model!A2:A3");
    helpers.select("Model!A1:B3");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 1 rows on Unpivot");
    expect(helpers.value("Unpivot!A2")).toBe("Group1");
    expect(helpers.value("Unpivot!C2")).toBe(10);
    // The 20 under the merge has no row key of its own and is never written.
    expect(helpers.value("Unpivot!A3")).toBe("");
  });
});

// ---------------------------------------------------------------------------
// A 31-character source sheet name (Excel's own maximum): unpivot never
// writes a sheet-qualified formula, so nothing here should even notice.
// ---------------------------------------------------------------------------

describe("a 31-character source sheet name", () => {
  it("reads and unpivots the selection exactly as it would from a short name", async () => {
    const longName = "ABCDEFGHIJKLMNOPQRSTUVWXYZ12345";
    expect(longName).toHaveLength(31);
    await boot([longName]);
    helpers.seed(`${longName}!A1`, [
      ["", "North"],
      ["Q1", 10],
    ]);
    helpers.select(`${longName}!A1:B2`);

    expect(await smt.unpivotSelection()).toBe("Unpivot: 1 rows on Unpivot");
    expect(helpers.value("Unpivot!C2")).toBe(10);
  });
});

// ---------------------------------------------------------------------------
// Pressed twice: unpivot never captures pls,fix Undo (it only ever adds a
// sheet, never overwrites a cell), so two presses make two sheets and leave
// nothing for Undo to do.
// ---------------------------------------------------------------------------

describe("pressed twice in a row", () => {
  it("makes a second sheet rather than reusing or refusing the first, and leaves the source untouched", async () => {
    helpers.seed("Model!A1", [
      ["", "North"],
      ["Q1", 10],
    ]);
    helpers.select("Model!A1:B2");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 1 rows on Unpivot");
    expect(await smt.unpivotSelection()).toBe("Unpivot: 1 rows on Unpivot 2");
    expect(helpers.value("Unpivot!C2")).toBe(10);
    expect(helpers.value("Unpivot 2!C2")).toBe(10);
    expect(helpers.value("Model!B2")).toBe(10);

    expect(await rejects(() => smt.undoLastAction())).toBe(
      "There is no pls,fix action to undo yet.",
    );
  });
});
