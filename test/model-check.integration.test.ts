// The model check against the strict fake host: that each kind is found where
// it really sits, that the pass reads without writing, and that a sheet too
// large to read is named rather than quietly left out.

import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type { CheckKind, Finding } from "../src/model-check";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(sheets: string[] = ["Model", "Data", "Scratch"]) {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
}

function of(findings: Finding[], kind: CheckKind): Finding[] {
  return findings.filter((entry) => entry.kind === kind);
}

function where(findings: Finding[], kind: CheckKind): string[] {
  return of(findings, kind).map(
    (entry) => `${entry.sheet ?? "-"}!${entry.ref ?? "-"}`,
  );
}

// One workbook carrying one of everything the check looks for.
function seedTheLot(): void {
  // A cell whose formula stopped resolving: the value IS the error text.
  helpers.seed("Model!B2", [[{ value: "#REF!", formula: "=Data!#REF!" }]]);
  // A number typed into a formula.
  helpers.seed("Model!B4", [[{ value: 11, formula: "=B3*1.1" }]]);
  // A filled row with one cell that was edited by hand. R1C1 is what says so:
  // a copied formula reads identically in every cell it was filled into.
  helpers.seed("Data!B2", [
    [
      { formula: "=A2", r1c1: "=RC[-1]", value: 4 },
      { formula: "=B2", r1c1: "=RC[-1]", value: 4 },
      { formula: "=C3", r1c1: "=R[1]C[-1]", value: 9 },
    ],
  ]);
  // A volatile lookup, written without a bare number so it is one finding.
  helpers.seed("Data!B5", [[{ value: 3, formula: "=OFFSET(A1,B1,C1)" }]]);
  // A formula pointing at another workbook.
  helpers.seed("Data!B7", [
    [{ value: 5, formula: "=[Budget.xlsx]Model!$B$4" }],
  ]);
  helpers.addName("Revenue", "=Model!$A$1");
  helpers.addName("Costs", "=Model!#REF!");
  // A custom style nothing wears. Excel's own styles come with the workbook.
  helpers.addStyle("Header 2");
}

beforeEach(async () => {
  await boot();
  helpers.sheet("Scratch").visibility = "Hidden";
});

describe("model check", () => {
  it("finds one of every kind, on the sheet and cell it sits on", async () => {
    seedTheLot();

    const report = await smt.runModelCheck();
    const { findings } = report;

    expect(where(findings, "formulaError")).toEqual(["Model!B2"]);
    expect(where(findings, "hardcodeInFormula")).toEqual(["Model!B4"]);
    expect(where(findings, "inconsistentFormula")).toEqual(["Data!D2"]);
    expect(where(findings, "volatileFormula")).toEqual(["Data!B5"]);
    expect(where(findings, "externalLink")).toEqual(["Data!B7"]);
    expect(where(findings, "hiddenSheet")).toEqual(["Scratch!-"]);
    expect(where(findings, "brokenName")).toEqual(["-!-"]);
    expect(where(findings, "unusedStyle")).toEqual(["-!-"]);
    expect(report.truncated).toBe(false);
    expect(report.skipped).toEqual([]);
  });

  it("carries the detail a reviewer needs on each line", async () => {
    seedTheLot();

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "formulaError")[0]?.note).toBe("#REF!");
    expect(of(findings, "volatileFormula")[0]?.note).toBe(
      "OFFSET in =OFFSET(A1,B1,C1)",
    );
    expect(of(findings, "brokenName")[0]?.note).toBe("Costs");
    expect(of(findings, "unusedStyle")[0]?.note).toBe("Header 2");
    expect(of(findings, "hiddenSheet")[0]?.note).toBe("Hidden");
  });

  it("also finds a name broken on a sheet, noted as Sheet!Name", async () => {
    seedTheLot();
    helpers
      .sheet("Model")
      .names.push({ name: "TaxRate", formula: "=Model!#REF!", visible: true });
    const before = helpers.syncCount();

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "brokenName").map((entry) => entry.note)).toEqual([
      "Costs",
      "Model!TaxRate",
    ]);
    // Sheet-scoped names ride scanSheets's existing used-range-extent batch:
    // still 3 syncs there (plus the style scrubber's own 3), not a 4th here.
    expect(helpers.syncCount() - before).toBe(6);
  });

  it("lists a very hidden sheet as such, and activates nothing", async () => {
    helpers.sheet("Scratch").visibility = "VeryHidden";
    const activeBefore = workbook.activeSheetId;

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "hiddenSheet")[0]?.note).toBe(
      "Very hidden, and can only be shown outside Excel",
    );
    expect(workbook.activeSheetId).toBe(activeBefore);
  });

  it("reads the workbook without writing a cell", async () => {
    seedTheLot();
    const before = {
      Model: helpers.cellMap("Model"),
      Data: helpers.cellMap("Data"),
      Scratch: helpers.cellMap("Scratch"),
    };
    const selection = workbook.selection;

    await smt.runModelCheck();

    expect(helpers.cellMap("Model")).toEqual(before.Model);
    expect(helpers.cellMap("Data")).toEqual(before.Data);
    expect(helpers.cellMap("Scratch")).toEqual(before.Scratch);
    expect(workbook.selection).toEqual(selection);
    // Nothing was deleted either: the two scrubbers only ever reported.
    expect(workbook.names.map((entry) => entry.name)).toEqual([
      "Revenue",
      "Costs",
    ]);
    expect(workbook.styles.map((entry) => entry.name)).toContain("Header 2");
  });

  it("scans hidden sheets too", async () => {
    helpers.seed("Scratch!C3", [[{ value: 7, formula: "=C2*1.07" }]]);

    const { findings } = await smt.runModelCheck();

    expect(where(findings, "hardcodeInFormula")).toEqual(["Scratch!C3"]);
  });

  it("orders the findings by sheet, errors first, then reading order", async () => {
    seedTheLot();

    const { findings } = await smt.runModelCheck();

    expect(
      findings.map((entry) => `${entry.sheet ?? "Workbook"}:${entry.kind}`),
    ).toEqual([
      "Model:formulaError",
      "Model:hardcodeInFormula",
      "Data:inconsistentFormula",
      "Data:volatileFormula",
      "Data:externalLink",
      "Scratch:hiddenSheet",
      "Workbook:brokenName",
      "Workbook:unusedStyle",
    ]);
  });

  it("names the sheet it was too large to read and counts what it did read", async () => {
    helpers.seed("Model!A1", [[{ value: 2, formula: "=A2*1.5" }]]);
    helpers.seed("Data!A1", [
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);

    const report = await smt.runModelCheck(4, 100);

    expect(report.skipped).toEqual(["Data"]);
    expect(report.scanned).toEqual({ sheets: 2, cells: 1 });
    expect(where(report.findings, "hardcodeInFormula")).toEqual(["Model!A1"]);
  });

  it("says nothing to flag on a workbook with nothing in it", async () => {
    await boot(["Only"]);
    const pure = await import("../src/model-check");

    const report = await smt.runModelCheck();

    expect(report.findings).toEqual([]);
    expect(pure.summarize(report)).toBe(
      "Model check: nothing to flag on 1 sheet",
    );
  });

  it("keeps a style that only a hidden sheet wears out of the report", async () => {
    helpers.addStyle("Header 2");
    helpers.seed("Scratch!A1", [["kept"]]);
    helpers.setStyle("Scratch!A1", "Header 2");

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "unusedStyle")).toEqual([]);
  });
});
