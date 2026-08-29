// Super Find against the strict fake host: what one pass over a workbook finds
// in cells, defined names and sheet names, which sheets it refuses to read, and
// where a result jumps to.

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

const LOOSE = { matchCase: false, inFormulas: false };
const FORMULAS = { matchCase: false, inFormulas: true };

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

beforeEach(async () => {
  await boot();
});

describe("find in workbook", () => {
  it("finds cells on every sheet, hidden ones included", async () => {
    helpers.seed("Model!B2", [["Revenue"]]);
    helpers.seed("Data!C3", [["revenue growth"]]);
    helpers.sheet("Data").visibility = "Hidden";

    expect(await smt.findInWorkbook("revenue", LOOSE)).toEqual({
      hits: [
        { kind: "cell", sheet: "Model", address: "B2", text: "Revenue" },
        { kind: "cell", sheet: "Data", address: "C3", text: "revenue growth" },
      ],
      skippedSheets: [],
    });
  });

  it("addresses a hit from the sheet's origin, not the used range's", async () => {
    helpers.seed("Model!D9", [
      ["Costs", "Total"],
      [10, 20],
    ]);

    const { hits } = await smt.findInWorkbook("Total", LOOSE);
    expect(hits).toEqual([
      { kind: "cell", sheet: "Model", address: "E9", text: "Total" },
    ]);
  });

  it("searches formulas only when asked, and shows the formula", async () => {
    helpers.seed("Model!B2", [[{ value: 84, formula: "=Drivers!B2*2" }]]);

    expect((await smt.findInWorkbook("Drivers", LOOSE)).hits).toEqual([]);
    expect((await smt.findInWorkbook("Drivers", FORMULAS)).hits).toEqual([
      { kind: "cell", sheet: "Model", address: "B2", text: "=Drivers!B2*2" },
    ]);
  });

  it("honours match case", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);

    const strict = { matchCase: true, inFormulas: false };
    expect((await smt.findInWorkbook("revenue", strict)).hits).toEqual([]);
    expect((await smt.findInWorkbook("Revenue", strict)).hits).toHaveLength(1);
  });

  it("finds a defined name by its name and by its formula", async () => {
    helpers.addName("Revenue_growth", "=Model!$A$1");

    expect((await smt.findInWorkbook("growth", LOOSE)).hits).toEqual([
      {
        kind: "name",
        sheet: "",
        address: "Revenue_growth",
        text: "=Model!$A$1",
      },
    ]);
    expect((await smt.findInWorkbook("$A$1", LOOSE)).hits).toHaveLength(1);
  });

  it("never surfaces a link anchor among the names", async () => {
    helpers.addName("SMT_LINK_5f3a91c2", "=Model!$A$1");
    helpers.addName("Anchor_note", "=Model!$A$1");

    expect((await smt.findInWorkbook("$A$1", LOOSE)).hits).toEqual([
      { kind: "name", sheet: "", address: "Anchor_note", text: "=Model!$A$1" },
    ]);
  });

  it("finds a sheet by its own name", async () => {
    await boot(["Model", "Data build"]);

    expect((await smt.findInWorkbook("build", LOOSE)).hits).toEqual([
      { kind: "sheet", sheet: "Data build", address: "A1", text: "Data build" },
    ]);
  });

  it("lists names first, then each sheet's name before its cells", async () => {
    await boot(["Alpha", "Beta"]);
    helpers.addName("Alpha_total", "=Alpha!$A$1");
    helpers.seed("Alpha!A1", [["Alpha opening"]]);
    helpers.seed("Beta!A1", [["Alpha closing"]]);

    const { hits } = await smt.findInWorkbook("Alpha", LOOSE);
    expect(hits.map((hit) => `${hit.kind}:${hit.sheet}${hit.address}`)).toEqual(
      ["name:Alpha_total", "sheet:AlphaA1", "cell:AlphaA1", "cell:BetaA1"],
    );
  });

  it("skips a sheet whose used range is past the cap and says which", async () => {
    helpers.seed("Model!A1", [["Total"]]);
    helpers.seed("Data!A1", [
      ["Total", "Total"],
      ["Total", "Total"],
    ]);

    expect(
      await smt.findInWorkbook("Total", { ...LOOSE, maxCells: 3 }),
    ).toEqual({
      hits: [{ kind: "cell", sheet: "Model", address: "A1", text: "Total" }],
      skippedSheets: ["Data"],
    });
  });

  it("finds nothing in an empty workbook without reading a null range", async () => {
    expect(await smt.findInWorkbook("anything", LOOSE)).toEqual({
      hits: [],
      skippedSheets: [],
    });
  });
});

describe("jump to a hit", () => {
  it("activates the sheet and selects the cell", async () => {
    helpers.seed("Data!B7", [["Target"]]);
    const { hits } = await smt.findInWorkbook("Target", LOOSE);

    await smt.jumpToHit(hits[0]!);

    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
    expect(workbook.selection).toEqual({
      sheetId: helpers.sheet("Data").id,
      rect: { row: 6, col: 1, rowCount: 1, colCount: 1 },
    });
  });

  it("selects the range a defined name points at", async () => {
    helpers.addName("Assumptions", "=Data!$C$3:$D$4");
    const { hits } = await smt.findInWorkbook("Assumptions", LOOSE);

    await smt.jumpToHit(hits[0]!);

    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
    expect(workbook.selection.rect).toEqual({
      row: 2,
      col: 2,
      rowCount: 2,
      colCount: 2,
    });
  });

  it("goes to A1 of a sheet found by name", async () => {
    await smt.activateSheet("Model");
    const { hits } = await smt.findInWorkbook("Data", LOOSE);

    await smt.jumpToHit(hits[0]!);

    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
    expect(workbook.selection.rect).toEqual({
      row: 0,
      col: 0,
      rowCount: 1,
      colCount: 1,
    });
  });

  it("refuses a hidden sheet rather than letting Excel throw", async () => {
    helpers.seed("Data!B7", [["Target"]]);
    const { hits } = await smt.findInWorkbook("Target", LOOSE);
    helpers.sheet("Data").visibility = "Hidden";

    expect(await rejects(() => smt.jumpToHit(hits[0]!))).toBe(
      "Data is hidden, so there is nowhere to jump.",
    );
    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
  });
});
