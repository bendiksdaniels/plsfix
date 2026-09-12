// "Select consistent region" against the strict fake host: the sheet's used
// range is read once, the rectangle of identical R1C1 formulas around the
// active cell becomes the selection, and nothing is written. Holds the flow to
// its four syncs and to the universality rows every Excel tool answers.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  formatA1,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
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

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

const GROWTH = "=RC[-1]*1.05";
const OTHER = "=RC[-1]*1.5";

// A formula cell the way Excel serves one: a value to make the cell used, and
// the same R1C1 string in both formula readings.
function formula(text: string): {
  formula: string;
  r1c1: string;
  value: number;
} {
  return { formula: text, r1c1: text, value: 1 };
}

function selected(): string {
  return formatA1(workbook.selection.rect);
}

describe("selectConsistentRegion", () => {
  it("grows a filled row from the active cell and counts it", async () => {
    helpers.seed("Model!B2", [
      ["Label", formula(GROWTH), formula(GROWTH), formula(GROWTH)],
    ]);
    helpers.select("Model!D2");

    expect(await smt.selectConsistentRegion()).toBe(
      "3 cells share this formula",
    );
    expect(selected()).toBe("C2:E2");
  });

  it("grows a block in both directions", async () => {
    helpers.seed("Model!B2", [
      [formula(GROWTH), formula(GROWTH), formula(GROWTH)],
      [formula(GROWTH), formula(GROWTH), formula(GROWTH)],
    ]);
    helpers.select("Model!C3");

    expect(await smt.selectConsistentRegion()).toBe(
      "6 cells share this formula",
    );
    expect(selected()).toBe("B2:D3");
  });

  it("starts from the active cell, not the selection's corner", async () => {
    helpers.seed("Model!A1", [
      [formula(OTHER), 12],
      [3, formula(GROWTH)],
      [null, formula(GROWTH)],
    ]);
    helpers.select("Model!A1:B3");
    helpers.setActiveCell("Model!B2");

    expect(await smt.selectConsistentRegion()).toBe(
      "2 cells share this formula",
    );
    expect(selected()).toBe("B2:B3");
  });

  it("says so when only the active cell holds that formula", async () => {
    helpers.seed("Model!A1", [[formula(OTHER), formula(GROWTH), 4]]);
    helpers.select("Model!B1");

    expect(await smt.selectConsistentRegion()).toBe(
      "Only the active cell has this formula.",
    );
    expect(selected()).toBe("B1");
  });

  it("says so when the active cell holds no formula", async () => {
    helpers.seed("Model!A1", [["Revenue", 120]]);
    helpers.select("Model!B1");

    expect(await smt.selectConsistentRegion()).toBe(
      "The active cell has no formula.",
    );
  });

  it("says so on an empty selection of an empty sheet", async () => {
    helpers.select("Model!C7");
    expect(await smt.selectConsistentRegion()).toBe(
      "The active cell has no formula.",
    );
  });

  it("says so when the active cell sits outside the used range", async () => {
    helpers.seed("Model!A1", [[formula(GROWTH)]]);
    helpers.select("Model!H20");

    expect(await smt.selectConsistentRegion()).toBe(
      "The active cell has no formula.",
    );
  });

  it("refuses a ctrl-clicked selection with the staged sentence", async () => {
    helpers.seed("Model!A1", [[formula(GROWTH), formula(GROWTH)]]);
    helpers.selectAreas(["Model!A1", "Model!C3"]);

    expect(await rejects(() => smt.selectConsistentRegion())).toBe(
      "Consistent region: select a single range",
    );
  });

  it("refuses a used range past the overlay's scan cap", async () => {
    helpers.seed("Model!A1", [[formula(GROWTH)]]);
    helpers.seed("Model!A6000", [[42]]);
    helpers.select("Model!A1");

    expect(await rejects(() => smt.selectConsistentRegion())).toBe(
      "The audit overlay supports up to 5,000 cells at once.",
    );
  });

  it("stops at a merged cell instead of throwing", async () => {
    // D2:E2 is one merged block, so only D2 carries the formula and E2 reads
    // blank - exactly what Excel serves for the cells a merge swallowed.
    helpers.seed("Model!B2", [
      [formula(GROWTH), formula(GROWTH), formula(GROWTH)],
    ]);
    helpers.merge("Model!D2:E2");
    helpers.select("Model!B2");

    expect(await smt.selectConsistentRegion()).toBe(
      "3 cells share this formula",
    );
    expect(selected()).toBe("B2:D2");
  });

  it("still answers on a protected sheet: it reads and selects, never writes", async () => {
    helpers.seed("Model!B2", [[formula(GROWTH), formula(GROWTH)]]);
    helpers.protectSheet("Model");
    helpers.select("Model!B2");

    expect(await smt.selectConsistentRegion()).toBe(
      "2 cells share this formula",
    );
    expect(selected()).toBe("B2:C2");
  });

  it("reads the grid once and selects once: four syncs", async () => {
    helpers.seed("Model!B2", [[formula(GROWTH), formula(GROWTH)]]);
    helpers.select("Model!B2");

    const before = helpers.syncCount();
    await smt.selectConsistentRegion();
    expect(helpers.syncCount() - before).toBe(4);
  });
});
