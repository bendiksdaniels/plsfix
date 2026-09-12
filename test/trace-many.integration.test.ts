// "Precedents of selection" against the strict fake host: every formula cell
// of one selected block asks Excel for its direct precedents in a single
// batch, the answers are grouped by source cell and unioned, and the
// same-sheet part of that union becomes the selection. Read-only.

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

function formula(text: string): {
  formula: string;
  r1c1: string;
  value: number;
} {
  return { formula: text, r1c1: text, value: 1 };
}

function selected(): string {
  return workbook.selectionAreas.length > 0
    ? workbook.selectionAreas.map((area) => formatA1(area.rect)).join(",")
    : formatA1(workbook.selection.rect);
}

// Three formula cells side by side, the shape of a model row a reviewer picks.
function seedRow(): void {
  helpers.seed("Model!B2", [
    [formula("=RC[-1]"), formula("=RC[-1]"), formula("=RC[-1]")],
  ]);
  helpers.select("Model!B2:D2");
}

describe("tracePrecedentsOfSelection", () => {
  it("groups the answers by source cell and unions the areas", async () => {
    seedRow();
    helpers.setPrecedents("Model!B2", [
      { address: "Model!A1:A3", cellCount: 3 },
    ]);
    helpers.setPrecedents("Model!C2", [
      { address: "Model!A1:A3", cellCount: 3 },
      { address: "Model!F9", cellCount: 1 },
    ]);
    helpers.setPrecedents("Model!D2", [{ address: "Model!F9", cellCount: 1 }]);

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.origin).toBe("Model!B2:D2");
    expect(result.formulaCells).toBe(3);
    expect(result.skipped).toBe(0);
    expect(result.groups).toEqual([
      {
        cell: "B2",
        areas: [{ sheet: "Model", address: "A1:A3", cellCount: 3 }],
      },
      {
        cell: "C2",
        areas: [
          { sheet: "Model", address: "A1:A3", cellCount: 3 },
          { sheet: "Model", address: "F9", cellCount: 1 },
        ],
      },
      { cell: "D2", areas: [{ sheet: "Model", address: "F9", cellCount: 1 }] },
    ]);
    // The same area under two source cells is one precedent, not two.
    expect(result.areas).toEqual([
      { sheet: "Model", address: "A1:A3", cellCount: 3 },
      { sheet: "Model", address: "F9", cellCount: 1 },
    ]);
  });

  it("selects the same-sheet union and only lists the other sheets", async () => {
    seedRow();
    helpers.setPrecedents("Model!B2", [{ address: "Model!A1", cellCount: 1 }]);
    helpers.setPrecedents("Model!C2", [
      { address: "Data!C5:C6", cellCount: 2 },
    ]);
    helpers.setPrecedents("Model!D2", [{ address: "Model!A2", cellCount: 1 }]);

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.areas.map((area) => `${area.sheet}!${area.address}`)).toEqual(
      ["Model!A1", "Data!C5:C6", "Model!A2"],
    );
    expect(selected()).toBe("A1,A2");
    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
  });

  it("skips the cells that hold no formula and counts them", async () => {
    helpers.seed("Model!B2", [[formula("=RC[-1]"), 120, "Total"]]);
    helpers.select("Model!B2:D2");
    helpers.setPrecedents("Model!B2", [{ address: "Model!A2", cellCount: 1 }]);

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.formulaCells).toBe(1);
    expect(result.skipped).toBe(2);
    expect(result.groups.map((group) => group.cell)).toEqual(["B2"]);
  });

  it("answers an empty selection without asking Excel anything", async () => {
    helpers.select("Model!G7:H8");

    const before = helpers.syncCount();
    const result = await smt.tracePrecedentsOfSelection();

    expect(result).toEqual({
      origin: "Model!G7:H8",
      groups: [],
      areas: [],
      formulaCells: 0,
      skipped: 4,
    });
    // The area count and the cell count, the grid, and then nothing.
    expect(helpers.syncCount() - before).toBe(3);
  });

  it("reads Excel's ItemNotFound as a source cell that reads from nothing", async () => {
    seedRow();
    helpers.setPrecedents("Model!B2", [{ address: "Model!A1", cellCount: 1 }]);
    helpers.setPrecedents("Model!D2", [{ address: "Model!A2", cellCount: 1 }]);

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.groups).toEqual([
      { cell: "B2", areas: [{ sheet: "Model", address: "A1", cellCount: 1 }] },
      { cell: "C2", areas: [] },
      { cell: "D2", areas: [{ sheet: "Model", address: "A2", cellCount: 1 }] },
    ]);
    expect(result.areas).toHaveLength(2);
  });

  it("refuses more cells than one batch may ask about", async () => {
    helpers.select("Model!A1:A51");
    expect(await rejects(() => smt.tracePrecedentsOfSelection())).toBe(
      "Precedents of selection handles up to 50 cells at once.",
    );
  });

  it("refuses a ctrl-clicked selection with the staged sentence", async () => {
    helpers.selectAreas(["Model!A1", "Model!C3"]);
    expect(await rejects(() => smt.tracePrecedentsOfSelection())).toBe(
      "Precedents of selection: select a single range",
    );
  });

  it("says so on a host too old to trace", async () => {
    seedRow();
    helpers.setSupported((_set, version) => version !== "1.12");
    expect(await rejects(() => smt.tracePrecedentsOfSelection())).toBe(
      "Tracing needs a newer Excel build.",
    );
  });

  it("reads the cells a merge swallowed as blank instead of throwing", async () => {
    helpers.seed("Model!B2", [[formula("=RC[-1]")]]);
    helpers.merge("Model!B2:C2");
    helpers.select("Model!B2:C2");
    helpers.setPrecedents("Model!B2", [{ address: "Model!A2", cellCount: 1 }]);

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.formulaCells).toBe(1);
    expect(result.skipped).toBe(1);
    expect(selected()).toBe("A2");
  });

  it("still answers on a protected sheet: it reads and selects, never writes", async () => {
    seedRow();
    helpers.protectSheet("Model");
    helpers.setPrecedents("Model!B2", [{ address: "Model!A1", cellCount: 1 }]);
    helpers.setPrecedents("Model!C2", "itemNotFound");
    helpers.setPrecedents("Model!D2", "itemNotFound");

    const result = await smt.tracePrecedentsOfSelection();

    expect(result.areas).toHaveLength(1);
    expect(selected()).toBe("A1");
  });

  it("one load for the formulas, one for the precedents, one select: five syncs", async () => {
    seedRow();
    for (const cell of ["Model!B2", "Model!C2", "Model!D2"]) {
      helpers.setPrecedents(cell, [{ address: "Model!A1", cellCount: 1 }]);
    }

    const before = helpers.syncCount();
    await smt.tracePrecedentsOfSelection();
    expect(helpers.syncCount() - before).toBe(5);
  });
});
