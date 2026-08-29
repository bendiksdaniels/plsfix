// Prepare for sharing against the strict fake host: where the workbook is left
// standing, what the pass refuses to touch, and what it reports back.

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

const A1 = { row: 0, col: 0, rowCount: 1, colCount: 1 };

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(
  sheets: string[] = ["Model", "Data", "Scratch"],
): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
}

// The fake keeps one selection, not one per sheet, so the trail of selections
// the run made is what says every visible sheet went back to A1.
function watchSelections(book: FakeWorkbook): () => string[] {
  const trail: string[] = [];
  let current = book.selection;
  Object.defineProperty(book, "selection", {
    configurable: true,
    get: () => current,
    set: (value: FakeWorkbook["selection"]) => {
      current = value;
      trail.push(
        `${book.find(value.sheetId)?.name ?? "?"}!${formatA1(value.rect)}`,
      );
    },
  });
  return () => trail;
}

function labels(
  report: ExcelModule.ShareResult["report"],
  kind: string,
): string[] {
  return report
    .filter((issue) => issue.kind === kind)
    .map((issue) => issue.label);
}

beforeEach(async () => {
  await boot();
  helpers.sheet("Scratch").visibility = "Hidden";
});

describe("prepare for sharing", () => {
  it("selects A1 on every visible sheet and ends on the first", async () => {
    helpers.seed("Data!C5", [["Working here"]]);
    helpers.select("Data!C5");
    const trail = watchSelections(workbook);

    const result = await smt.prepareForSharing();

    expect(trail()).toEqual(["Model!A1", "Data!A1"]);
    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
    expect(workbook.selection.rect).toEqual(A1);
    expect(result.touchedSheets).toBe(2);
  });

  it("leaves the hidden sheet alone and reports it", async () => {
    helpers.seed("Scratch!B2", [["Old working"]]);
    const before = helpers.cellMap("Scratch");
    const trail = watchSelections(workbook);

    const { report } = await smt.prepareForSharing();

    expect(trail()).not.toContain("Scratch!A1");
    expect(helpers.sheet("Scratch").visibility).toBe("Hidden");
    expect(helpers.cellMap("Scratch")).toEqual(before);
    expect(labels(report, "hiddenSheet")).toEqual(["Scratch"]);
  });

  it("reports a formula pointing at another workbook, hidden sheets included", async () => {
    helpers.seed("Model!B4", [
      [{ value: 12, formula: "=[Budget.xlsx]Model!$B$4" }],
    ]);
    helpers.seed("Scratch!A2", [[{ value: 3, formula: "=[Old.xlsx]S!$A$1" }]]);
    helpers.seed("Data!A1", [[{ value: 5, formula: "=Model!B4" }]]);

    const { report } = await smt.prepareForSharing();

    expect(labels(report, "externalLink")).toEqual([
      "Model!B4: =[Budget.xlsx]Model!$B$4",
      "Scratch!A2: =[Old.xlsx]S!$A$1",
    ]);
  });

  it("reports broken names without deleting one", async () => {
    helpers.addName("Revenue", "=Model!$A$1");
    helpers.addName("Costs", "=Model!#REF!");
    helpers.addName("PLSFIX_LINK_5f3a91c2", "=Model!#REF!");

    const { report } = await smt.prepareForSharing();

    expect(labels(report, "brokenName")).toEqual(["Costs"]);
    expect(workbook.names.map((entry) => entry.name)).toEqual([
      "Revenue",
      "Costs",
      "PLSFIX_LINK_5f3a91c2",
    ]);
  });

  it("reports a sheet whose used range is too large to scan", async () => {
    helpers.seed("Model!A1", [[{ value: 1, formula: "=[Budget.xlsx]S!$A$1" }]]);
    helpers.seed("Data!A1", [
      ["a", "b"],
      ["c", "d"],
    ]);

    const { report } = await smt.prepareForSharing({ maxCells: 3 });

    expect(labels(report, "skippedSheet")).toEqual(["Data"]);
    expect(labels(report, "externalLink")).toEqual([
      "Model!A1: =[Budget.xlsx]S!$A$1",
    ]);
  });

  it("stops at the scan cap and names the sheets it did not read", async () => {
    await boot(["Model", "Data", "Notes"]);
    for (const sheet of ["Model", "Data", "Notes"]) {
      helpers.seed(`${sheet}!A1`, [
        ["a", "b"],
        [{ value: 1, formula: "=[Budget.xlsx]S!$A$1" }, "d"],
      ]);
    }

    const { report } = await smt.prepareForSharing({
      maxCells: 4,
      maxTotalCells: 8,
    });

    // The link on the sheet that was not read is missing from the report, and
    // the report says so rather than reading as "nothing there".
    expect(labels(report, "skippedSheet")).toEqual(["Notes"]);
    expect(labels(report, "externalLink")).toEqual([
      "Model!A2: =[Budget.xlsx]S!$A$1",
      "Data!A2: =[Budget.xlsx]S!$A$1",
    ]);
  });

  // The three this release introduced: the add-in's own functions, the paint an
  // overlay leaves in the file, and the tokens the link registry carries.
  it("reports the cells only this add-in can evaluate", async () => {
    helpers.seed("Model!B4", [
      [{ value: 12, formula: "=PLSFIX.ROUND(B3,-3)" }],
      [{ value: 12, formula: "=_xlfn.PLSFIX.ROUNDSUM(B1:B2,0)" }],
      [{ value: 12, formula: "=ROUND(B3,-3)" }],
    ]);

    const { report } = await smt.prepareForSharing();

    expect(labels(report, "addinFormula")).toEqual([
      "Model!B4: =PLSFIX.ROUND(B3,-3)",
      "Model!B5: =_xlfn.PLSFIX.ROUNDSUM(B1:B2,0)",
    ]);
  });

  it("reports an overlay still painted over the model", async () => {
    const formula = { formula: "=R[1]C", r1c1: "=R[1]C", value: 1 };
    helpers.seed("Model!A1", [
      [formula, formula],
      [formula, formula],
    ]);
    helpers.select("Model!A1:B2");
    expect(await smt.toggleAuditOverlay()).toBe(true);

    const { report } = await smt.prepareForSharing();

    expect(labels(report, "overlayPainted")).toEqual([
      "The audit overlay is still painted over cells; the reader has no add-in to clear it",
    ]);

    // The pass left the workbook on A1, so the overlay is toggled off from the
    // block it painted, and the next report has nothing to say about it.
    helpers.select("Model!A1:B2");
    expect(await smt.toggleAuditOverlay()).toBe(false);
    const off = await smt.prepareForSharing();
    expect(labels(off.report, "overlayPainted")).toEqual([]);
  });

  it("reports the link registry as tokens to break, and deletes nothing", async () => {
    const pure = await import("../src/share");
    const clean = await smt.prepareForSharing();
    expect(labels(clean.report, "linkTokens")).toEqual([]);

    // What the Links tab leaves behind: the registry setting, tokens and all.
    helpers.setSetting(
      "PLSFIX_LINKS",
      JSON.stringify({
        v: 1,
        links: [
          {
            id: "a".repeat(32),
            kind: "range",
            anchor: "PLSFIX_LINK_aaaaaaaa",
            label: "Model!B4:F5",
            token: "t".repeat(43),
            createdAt: "2026-08-29T10:00:00.000Z",
            lastPushedAt: null,
            rev: 1,
          },
        ],
      }),
    );

    const { report } = await smt.prepareForSharing();

    expect(labels(report, "linkTokens")).toEqual([pure.LINK_TOKENS_LABEL]);
    expect(helpers.setting("PLSFIX_LINKS")).not.toBe("");
  });

  it("reports autocolor on edit while it is running", async () => {
    const clean = await smt.prepareForSharing();
    expect(labels(clean.report, "autocolorOnEdit")).toEqual([]);

    await smt.setAutocolorOnEdit(true);
    const live = await smt.prepareForSharing();

    expect(labels(live.report, "autocolorOnEdit")).toHaveLength(1);
  });

  it("summarizes a clean workbook and a workbook with work left", async () => {
    const pure = await import("../src/share");

    helpers.sheet("Scratch").visibility = "Visible";
    const clean = await smt.prepareForSharing();
    expect(pure.summarizeShare(clean.report, clean.touchedSheets)).toBe(
      "3 sheets reset to A1; nothing else to fix",
    );

    helpers.sheet("Scratch").visibility = "Hidden";
    helpers.seed("Model!B4", [
      [{ value: 12, formula: "=[Budget.xlsx]Model!$B$4" }],
    ]);
    const left = await smt.prepareForSharing();
    expect(pure.summarizeShare(left.report, left.touchedSheets)).toBe(
      "2 sheets reset to A1; 1 hidden sheet, 1 external link left",
    );
  });

  it("still lands on A1 when the workbook has nothing on any sheet", async () => {
    await boot(["Only"]);

    const result = await smt.prepareForSharing();

    expect(result).toEqual({ report: [], touchedSheets: 1 });
    expect(workbook.selection).toEqual({
      sheetId: helpers.sheet("Only").id,
      rect: A1,
    });
  });
});
