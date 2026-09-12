// Comps stats against the strict fake host: the six live rows under a comps
// table, the formats and presets they wear, and every selection a modeller can
// hand it - multi-area, merged, protected, over the cap and empty.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let brand: typeof SettingsModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// A three-column comps table with a header row: labels, one numeric column, one
// text column, at A1:C4.
function seedComps(): void {
  helpers.seed("Model!A1", [
    ["Company", "EV/EBITDA", "Listed"],
    ["Alpha", 8.1, "yes"],
    ["Beta", 9.4, "yes"],
    ["Gamma", 7.2, "no"],
  ]);
  helpers.select("Model!A1:C4");
}

beforeEach(async () => {
  await boot();
});

describe("insertCompsStats", () => {
  it("writes six live rows one blank row under the block", async () => {
    seedComps();

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 3 rows",
    );

    // Row 5 stays blank, the block runs A6:C11.
    expect(helpers.value("Model!A5")).toBe("");
    expect(
      ["A6", "A7", "A8", "A9", "A10", "A11"].map((at) =>
        helpers.value(`Model!${at}`),
      ),
    ).toEqual([
      "Min",
      "25th percentile",
      "Median",
      "Mean",
      "75th percentile",
      "Max",
    ]);
    expect(
      ["B6", "B7", "B8", "B9", "B10", "B11"].map((at) =>
        helpers.formula(`Model!${at}`),
      ),
    ).toEqual([
      "=MIN(B2:B4)",
      "=PERCENTILE.INC(B2:B4,0.25)",
      "=MEDIAN(B2:B4)",
      "=AVERAGE(B2:B4)",
      "=PERCENTILE.INC(B2:B4,0.75)",
      "=MAX(B2:B4)",
    ]);
  });

  it("leaves the non-numeric columns empty", async () => {
    seedComps();
    await smt.insertCompsStats();

    expect(helpers.value("Model!C6")).toBe("");
    expect(helpers.value("Model!C11")).toBe("");
  });

  it("takes the number format off the last data row of each column", async () => {
    seedComps();
    helpers.setNumberFormat("Model!B2:B4", "0.0x");
    await smt.insertCompsStats();

    expect(helpers.numberFormat("Model!B6")).toBe("0.0x");
    expect(helpers.numberFormat("Model!B11")).toBe("0.0x");
    expect(helpers.numberFormat("Model!A6")).toBe("General");
  });

  it("gives the labels the plain look and the numbers the formula look", async () => {
    seedComps();
    await smt.insertCompsStats();

    const theme = brand.deriveTheme(brand.DEFAULT_SETTINGS);
    expect(helpers.font("Model!A6")).toMatchObject({
      bold: false,
      color: theme.formulaFont,
      size: 10,
    });
    expect(helpers.font("Model!B6").color).toBe(theme.formulaFont);
    // Nothing was painted over the column that stays empty.
    expect(helpers.font("Model!C6").size).not.toBe(10);
  });

  it("reads a header-free table and selects what it wrote", async () => {
    helpers.seed("Model!D10", [
      ["Alpha", 8.1],
      ["Beta", 9.4],
    ]);
    helpers.select("Model!D10:E11");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 2 rows",
    );
    expect(helpers.formula("Model!E13")).toBe("=MIN(E10:E11)");
    expect(helpers.value("Model!D13")).toBe("Min");
  });

  it("counts every numeric column of a wider table", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA", "P/E", "EV/Sales"],
      ["Alpha", 8.1, 12, 1.4],
      ["Beta", 9.4, 14, 1.8],
    ]);
    helpers.select("Model!A1:D3");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 3 columns over 2 rows",
    );
    expect(helpers.formula("Model!D10")).toBe("=MAX(D2:D3)");
  });

  it("puts the block back on Undo", async () => {
    seedComps();
    await smt.insertCompsStats();
    await smt.undoLastAction();

    expect(helpers.value("Model!A6")).toBe("");
    expect(helpers.value("Model!B6")).toBe("");
  });
});

describe("what Comps stats refuses", () => {
  it("refuses when the six target rows are not empty", async () => {
    seedComps();
    helpers.seed("Model!B8", [["do not overwrite"]]);

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats need six empty rows under the block.",
    );
    expect(helpers.value("Model!A6")).toBe("");
  });

  it("refuses a table with fewer than two data rows", async () => {
    helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
    ]);
    helpers.select("Model!A1:B2");

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats need a table with at least two data rows.",
    );
  });

  it("refuses a table with no numeric column", async () => {
    helpers.seed("Model!A1", [
      ["Company", "Sector"],
      ["Alpha", "banks"],
      ["Beta", "banks"],
    ]);
    helpers.select("Model!A1:B3");

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats need at least one column of numbers.",
    );
  });

  it("refuses an empty selection the same way", async () => {
    helpers.select("Model!C3:E6");

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats need at least one column of numbers.",
    );
  });

  it("refuses a ctrl-clicked selection", async () => {
    seedComps();
    helpers.selectAreas(["Model!A1:C4", "Model!E1:F2"]);

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats: select a single range",
    );
  });

  it("refuses a selection over the cell cap, before any grid is read", async () => {
    helpers.select("Model!A:C");

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats supports up to 5,000 selected cells at once.",
    );
  });

  it("names the protected sheet instead of Excel's own string", async () => {
    seedComps();
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.insertCompsStats())).toBe(
      "Comps stats: this sheet is protected, nothing was changed",
    );
    expect(helpers.value("Model!A6")).toBe("");
  });
});

describe("Comps stats on a merged selection", () => {
  it("reads a merged heading band as the header row", async () => {
    helpers.seed("Model!A1", [["Trading comparables"]]);
    helpers.merge("Model!A1:B1");
    helpers.seed("Model!A2", [
      ["Alpha", 8.1],
      ["Beta", 9.4],
    ]);
    helpers.select("Model!A1:B3");

    expect(await smt.insertCompsStats()).toBe(
      "Comps stats written: 1 column over 2 rows",
    );
    expect(helpers.formula("Model!B5")).toBe("=MIN(B2:B3)");
  });
});
