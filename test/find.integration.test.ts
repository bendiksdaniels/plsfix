// Super Find against the strict fake host: what one pass over a workbook finds
// in cells, defined names, sheet names and comments, which sheets it refuses to
// read, what an old host cannot search, and where a result jumps to.

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
      commentsSkipped: false,
      scannedSheets: 2,
      sheetCap: 200_000,
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
    helpers.addName("PLSFIX_LINK_5f3a91c2", "=Model!$A$1");
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
      commentsSkipped: false,
      scannedSheets: 1,
      sheetCap: 3,
    });
  });

  // The per-sheet cap does not bound one request: sheets each just under it
  // add up to a payload the host refuses, with no stage and no sheet named.
  it("stops at the scan cap and names the sheets it did not read", async () => {
    await boot(["Model", "Data", "Notes"]);
    for (const sheet of ["Model", "Data", "Notes"]) {
      helpers.seed(`${sheet}!A1`, [
        ["Total", "Total"],
        ["Total", "Total"],
      ]);
    }

    expect(
      await smt.findInWorkbook("Total", {
        ...LOOSE,
        maxCells: 4,
        maxTotalCells: 8,
      }),
    ).toEqual({
      hits: [
        { kind: "cell", sheet: "Model", address: "A1", text: "Total" },
        { kind: "cell", sheet: "Model", address: "B1", text: "Total" },
        { kind: "cell", sheet: "Model", address: "A2", text: "Total" },
        { kind: "cell", sheet: "Model", address: "B2", text: "Total" },
        { kind: "cell", sheet: "Data", address: "A1", text: "Total" },
        { kind: "cell", sheet: "Data", address: "B1", text: "Total" },
        { kind: "cell", sheet: "Data", address: "A2", text: "Total" },
        { kind: "cell", sheet: "Data", address: "B2", text: "Total" },
      ],
      skippedSheets: ["Notes"],
      commentsSkipped: false,
      scannedSheets: 2,
      sheetCap: 4,
    });
  });

  // A deal model's largest sheet runs past the old selection-scan cap easily;
  // the default now matches the model check's own per-sheet cap so Find can
  // read it, and still names a sheet that runs past that.
  it("scans a sheet at the sheet-scan cap by default, skips one past it", async () => {
    helpers.seed("Model!A1", [["Total"]]);
    helpers.seed("Model!A50000", [[1]]);
    helpers.seed("Data!A1", [["Total revenue"]]);
    helpers.seed("Data!A250000", [[1]]);

    expect(await smt.findInWorkbook("Total", LOOSE)).toEqual({
      hits: [{ kind: "cell", sheet: "Model", address: "A1", text: "Total" }],
      skippedSheets: ["Data"],
      commentsSkipped: false,
      scannedSheets: 1,
      sheetCap: 200_000,
    });
  });

  it("finds nothing in an empty workbook without reading a null range", async () => {
    expect(await smt.findInWorkbook("anything", LOOSE)).toEqual({
      hits: [],
      skippedSheets: [],
      commentsSkipped: false,
      scannedSheets: 2,
      sheetCap: 200_000,
    });
  });
});

describe("find in comments", () => {
  it("lists a comment and its replies after the cells of their sheet", async () => {
    helpers.seed("Model!A1", [["Margin check"]]);
    helpers.addComment("Model!B2", "Margin looks light", "Anna Ozola", [
      { content: "Agreed, margin fixed", author: "Peteris Krumins" },
    ]);
    helpers.addComment("Data!C3", "Source: margin file", "Anna Ozola");

    expect((await smt.findInWorkbook("margin", LOOSE)).hits).toEqual([
      { kind: "cell", sheet: "Model", address: "A1", text: "Margin check" },
      {
        kind: "comment",
        sheet: "Model",
        address: "B2",
        text: "Margin looks light",
      },
      {
        kind: "comment",
        sheet: "Model",
        address: "B2",
        text: "(reply) Agreed, margin fixed",
      },
      {
        kind: "comment",
        sheet: "Data",
        address: "C3",
        text: "Source: margin file",
      },
    ]);
  });

  it("finds a comment by the name of whoever wrote it", async () => {
    helpers.addComment("Model!B2", "Looks fine", "Anna Ozola");

    expect((await smt.findInWorkbook("ozola", LOOSE)).hits).toEqual([
      { kind: "comment", sheet: "Model", address: "B2", text: "Looks fine" },
    ]);
  });

  it("still finds the comments on a sheet too large to read", async () => {
    helpers.seed("Data!A1", [
      ["Total", "Total"],
      ["Total", "Total"],
    ]);
    helpers.addComment("Data!D9", "Total is stale", "Anna Ozola");

    expect(
      await smt.findInWorkbook("Total", { ...LOOSE, maxCells: 3 }),
    ).toEqual({
      hits: [
        {
          kind: "comment",
          sheet: "Data",
          address: "D9",
          text: "Total is stale",
        },
      ],
      skippedSheets: ["Data"],
      commentsSkipped: false,
      scannedSheets: 1,
      sheetCap: 3,
    });
  });

  it("leaves comments out when the box is unticked", async () => {
    helpers.addComment("Model!B2", "Margin looks light", "Anna Ozola");

    expect(
      await smt.findInWorkbook("margin", { ...LOOSE, inComments: false }),
    ).toEqual({
      hits: [],
      skippedSheets: [],
      commentsSkipped: false,
      scannedSheets: 2,
      sheetCap: 200_000,
    });
  });

  it("says comments were skipped on a host below ExcelApi 1.10", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    helpers.addComment("Model!B2", "Margin looks light", "Anna Ozola");

    expect(await smt.findInWorkbook("margin", LOOSE)).toEqual({
      hits: [],
      skippedSheets: [],
      commentsSkipped: true,
      scannedSheets: 2,
      sheetCap: 200_000,
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

  it("selects the cell a comment hangs on", async () => {
    helpers.addComment("Data!B7", "Check this number", "Anna Ozola");
    const { hits } = await smt.findInWorkbook("Check this", LOOSE);

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

  // A name matches on its formula too, so a constant is a perfectly ordinary
  // hit; getRange() on one throws a bare host string with no flow in it.
  it("says which name has no range instead of throwing the host's message", async () => {
    helpers.addName("TaxRate", "=0.21");
    const { hits } = await smt.findInWorkbook("TaxRate", LOOSE);
    expect(hits[0]!.kind).toBe("name");

    expect(await rejects(() => smt.jumpToHit(hits[0]!))).toBe(
      'find: name "TaxRate" has no range',
    );
  });

  it("says the same for a name Excel has left on #REF!", async () => {
    helpers.addName("Costs", "=Data!$C$3");
    const { hits } = await smt.findInWorkbook("Costs", LOOSE);
    helpers.breakName("Costs");

    expect(await rejects(() => smt.jumpToHit(hits[0]!))).toBe(
      'find: name "Costs" has no range',
    );
  });
});
