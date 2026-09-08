// Audit suite for the sheet-chart, template, unpivot and reconcile tools: the
// hosts that refuse a chartex surface, the selections a modeller makes by
// clicking a column header, a chart anchored on rows the host cannot size, and
// the edge inputs each flow is documented to refuse.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], ...options });
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

function seedBridge(): void {
  helpers.seed("Model!A1", [["EBITDA bridge"]]);
  helpers.seed("Model!A2", [
    ["Opening", 100],
    ["Price", 20],
    ["Cost", -30],
    ["Closing", 90],
  ]);
  helpers.select("Model!A2:B5");
}

function seedDrivers(): void {
  helpers.seed("Model!A1", [["EBITDA sensitivity", 100]]);
  helpers.seed("Model!A2", [
    ["Driver", "Low", "High"],
    ["Volume", 90, 115],
    ["Price", 60, 140],
  ]);
  helpers.select("Model!A2:C4");
}

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------

describe("a clicked column header", () => {
  // The cap answers before a single value is read: two whole columns are two
  // million cells, and the shape rules only run after the read.
  it("is refused by the waterfall before its values are read", async () => {
    seedBridge();
    helpers.select("Model!A1:B1048576");

    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "Waterfall supports up to 5,000 selected cells at once.",
    );
  });

  it("is refused by the tornado before its values are read", async () => {
    seedDrivers();
    helpers.select("Model!A1:C1048576");

    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado supports up to 5,000 selected cells at once.",
    );
  });

  it("still names the point cap for a table under the cell cap", async () => {
    helpers.select("Model!A1:B101");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "A bridge chart supports up to 100 points.",
    );
  });
});

// ---------------------------------------------------------------------------

describe("the CAGR callout on what a modeller selects", () => {
  it("refuses a block and a single cell by name", async () => {
    helpers.seed("Model!A1", [
      [100, 110],
      [120, 130],
    ]);
    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "Select one row or column with at least two numbers.",
    );

    helpers.select("Model!A1");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "Select one row or column with at least two numbers.",
    );
  });

  it("refuses a line a compound rate cannot describe", async () => {
    helpers.seed("Model!A1", [[0, 110, 121]]);
    helpers.select("Model!A1:C1");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "A CAGR needs a positive start and end value.",
    );

    helpers.seed("Model!A5", [[-100, -110]]);
    helpers.select("Model!A5:B5");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "A CAGR needs a positive start and end value.",
    );
    expect(workbook.shapes).toHaveLength(0);
  });

  it("reads two points as one period", async () => {
    helpers.seed("Model!A1", [[100, 110]]);
    helpers.select("Model!A1:B1");
    expect(await smt.addCagrLabel()).toBe("CAGR +10.0% over 1 period");
  });
});

// ---------------------------------------------------------------------------

describe("unpivot on the tables a modeller has", () => {
  it("skips the columns a merged header leaves blank", async () => {
    // A merged header cell answers with its value in the top-left only, so the
    // column under the rest of the merge has no name to travel under.
    helpers.seed("Model!A1", [
      ["", "Half year", ""],
      ["Row", 1, 2],
    ]);
    helpers.merge("Model!B1:C1");
    helpers.select("Model!A1:C2");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 1 rows on Unpivot");
    expect(helpers.value("Unpivot!B2")).toBe("Half year");
  });

  it("refuses a single column by name", async () => {
    helpers.seed("Model!A1", [["Header"], ["Row"], ["Value"]]);
    helpers.select("Model!A1:A3");
    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: need a header row, a key column and one column of values",
    );
  });
});

// ---------------------------------------------------------------------------

describe("finding a combination in the awkward cases", () => {
  function seedValues(values: number[]): void {
    helpers.seed(
      "Model!A1",
      values.map((value) => [value]),
    );
    helpers.select(`Model!A1:A${String(values.length)}`);
  }

  it("solves a negative target", async () => {
    seedValues([-10, -25, 7, -15]);
    const result = await smt.reconcileSelection(-40, 0.001);
    expect(result.values.slice().sort((a, b) => a - b)).toEqual([-25, -15]);
    expect(result.sum).toBe(-40);
  });

  it("picks the fewest cells among duplicates", async () => {
    seedValues([20, 20, 20, 40]);
    const result = await smt.reconcileSelection(40, 0);
    expect(result.count).toBe(1);
    expect(result.addresses).toEqual(["A4"]);
  });

  it("takes the nearest sum when the tolerance is wider than the target", async () => {
    seedValues([12, 31]);
    const result = await smt.reconcileSelection(10, 1000);
    expect(result.values).toEqual([12]);
    expect(result.difference).toBe(2);
  });

  it("solves a block at the 34-value cap", async () => {
    const values = Array.from({ length: 34 }, (_unused, index) => index + 1);
    seedValues(values);
    // No pair of 1..34 reaches 69, so the fewest cells that do is three.
    const result = await smt.reconcileSelection(69, 0);
    expect(result.sum).toBe(69);
    expect(result.count).toBe(3);
  });
});
