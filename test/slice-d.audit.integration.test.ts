// Workbook tools against the strict fake host, the corners the feature suites
// leave out: the model check when the style table cannot be read, the contents
// sheet on a workbook with nothing visible left to list, the link anchors the
// name scrubber must never touch, a comment on a sheet whose name needs
// quoting, and what the share report says about a registry it cannot decode.

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

const LOOSE = { matchCase: false, inFormulas: false };

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

function of(findings: Finding[], kind: CheckKind): Finding[] {
  return findings.filter((entry) => entry.kind === kind);
}

// A used range past the style scrubber's own per-sheet cap without seeding a
// cell for every square of it: the corners are what the extent is read from.
function seedWideSheet(sheet: string): void {
  helpers.seed(`${sheet}!A1`, [["Corner"]]);
  helpers.seed(`${sheet}!BZ80`, [["Corner"]]);
}

beforeEach(async () => {
  await boot();
});

describe("model check when the style table cannot be read", () => {
  // The style scan is the heaviest read of the pass and the only one that
  // needs getCellProperties. A host refusing it must cost the reviewer the
  // style kind, not the other seven.
  it("still reports every other kind when getCellProperties is refused", async () => {
    helpers.addStyle("Header 2");
    helpers.seed("Model!B4", [[{ value: 11, formula: "=B3*1.1" }]]);
    helpers.addName("Costs", "=Model!#REF!");
    helpers.failNextCellProperties();

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "hardcodeInFormula")).toHaveLength(1);
    expect(of(findings, "brokenName")).toHaveLength(1);
    expect(of(findings, "unusedStyle")).toEqual([]);
  });

  // The same rule the scrubber already applies to a sheet it could not read:
  // a style could be worn anywhere on it, so the whole kind is dropped.
  it("drops the style kind when a sheet is too large for the style scan", async () => {
    helpers.addStyle("Header 2");
    seedWideSheet("Data");

    const { findings, skipped } = await smt.runModelCheck();

    expect(of(findings, "unusedStyle")).toEqual([]);
    // The pass itself read the wide sheet: only the style scan stopped short.
    expect(skipped).toEqual([]);
  });
});

describe("contents sheet", () => {
  it("clears its own rows when nothing visible is left to list", async () => {
    await smt.insertToc();
    expect(helpers.value("TOC!B3")).toBe("Model");

    helpers.sheet("Model").visibility = "Hidden";
    helpers.sheet("Data").visibility = "Hidden";
    await smt.insertToc();

    expect(helpers.value("TOC!A1")).toBe("pls,fix - Contents");
    expect(helpers.value("TOC!B3")).toBe("");
    expect(helpers.cell("TOC!B3").hyperlink).toBeNull();
  });

  it("brings its own sheet back into view when it was hidden", async () => {
    await smt.insertToc();
    helpers.sheet("TOC").visibility = "Hidden";

    await smt.insertToc();

    expect(helpers.sheet("TOC").visibility).toBe("Visible");
    expect(workbook.activeSheetId).toBe(helpers.sheet("TOC").id);
  });

  // Excel writes a sheet reference with its apostrophes doubled; a link that
  // does not would land nowhere.
  it("doubles an apostrophe in the sheet name it links to", async () => {
    await boot(["Bank's view"]);

    await smt.insertToc();

    expect(helpers.cell("TOC!B3").hyperlink).toEqual({
      documentReference: "'Bank''s view'!A1",
      textToDisplay: "Bank's view",
    });
  });
});

// The Map's invariant: a link's identity in Excel is the hidden name
// PLSFIX_LINK_<id8>. One whose rows were deleted is a #REF! name by design.
describe("the name scrubber and the link registry", () => {
  beforeEach(() => {
    helpers.addName("Costs", "=Model!#REF!");
    helpers.addName("PLSFIX_LINK_5f3a91c2", "=Model!#REF!");
    helpers.addName("PLSFIX_LINK_aaaaaaaa", "=Model!$A$1");
  });

  it("never lists a link anchor among the broken names", async () => {
    expect(await smt.listBrokenNames()).toEqual(["Costs"]);
  });

  it("never deletes a link anchor, broken or not", async () => {
    expect(await smt.deleteBrokenNames()).toBe(1);

    expect(workbook.names.map((entry) => entry.name)).toEqual([
      "PLSFIX_LINK_5f3a91c2",
      "PLSFIX_LINK_aaaaaaaa",
    ]);
  });
});

describe("super find, the corners", () => {
  it("finds a comment on a sheet whose name has to be quoted", async () => {
    await boot(["Model", "Data build"]);
    helpers.addComment("Data build!C3", "Margin looks light", "Anna Ozola");

    expect((await smt.findInWorkbook("margin", LOOSE)).hits).toEqual([
      {
        kind: "comment",
        sheet: "Data build",
        address: "C3",
        text: "Margin looks light",
      },
    ]);
  });

  it("refuses to jump to a name whose range sits on a hidden sheet", async () => {
    helpers.addName("Assumptions", "=Data!$C$3");
    const { hits } = await smt.findInWorkbook("Assumptions", LOOSE);
    helpers.sheet("Data").visibility = "Hidden";

    expect(await rejects(() => smt.jumpToHit(hits[0]!))).toBe(
      "Data is hidden, so there is nowhere to jump.",
    );
    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
  });

  // What a modeller types is what Excel shows: an error text, a number, a
  // Latvian word in whatever casing came to hand.
  it("finds an error value, a boolean and a word with diacritics", async () => {
    helpers.seed("Model!A1", [["#N/A", true, "Ieņēmumi"]]);

    expect((await smt.findInWorkbook("#N/A", LOOSE)).hits[0]?.address).toBe(
      "A1",
    );
    expect((await smt.findInWorkbook("true", LOOSE)).hits[0]?.text).toBe(
      "TRUE",
    );
    expect((await smt.findInWorkbook("IEŅĒMUMI", LOOSE)).hits[0]?.text).toBe(
      "Ieņēmumi",
    );
  });
});

describe("prepare for sharing, the corners", () => {
  it("lands on the first visible sheet when the first sheet is hidden", async () => {
    await boot(["Cover", "Model", "Data"]);
    helpers.sheet("Cover").visibility = "Hidden";

    const result = await smt.prepareForSharing();

    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
    expect(result.touchedSheets).toBe(2);
  });

  // A registry a newer build wrote still means the file carries tokens.
  it("counts a registry it cannot decode as tokens travelling with the file", async () => {
    const pure = await import("../src/share");
    helpers.setSetting("PLSFIX_LINKS", "{written by a newer build");

    const { report } = await smt.prepareForSharing();

    expect(
      report.filter((issue) => issue.kind === "linkTokens").map((i) => i.label),
    ).toEqual([pure.LINK_TOKENS_LABEL]);
  });

  it("says nothing about tokens for a registry holding no links", async () => {
    helpers.setSetting("PLSFIX_LINKS", JSON.stringify({ v: 1, links: [] }));

    const { report } = await smt.prepareForSharing();

    expect(report.filter((issue) => issue.kind === "linkTokens")).toEqual([]);
  });

  // The report is built from three reading phases and only then is a single
  // sheet touched: a host that refuses the read leaves the workbook exactly
  // where the modeller had it, not half tidied.
  it("moves nothing when the read the report is built from is refused", async () => {
    helpers.seed("Data!C5", [["Working here"]]);
    helpers.select("Data!C5");
    await smt.activateSheet("Data");
    const before = workbook.selection;
    helpers.failNextSync();

    expect(await rejects(() => smt.prepareForSharing())).toBe(
      "The sync failed.",
    );

    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
    expect(workbook.selection).toEqual(before);
  });
});

describe("the style scrubber, the corners", () => {
  // A back-room sheet nobody can open from Excel's UI still dresses its cells.
  it("counts a style worn only on a very hidden sheet as in use", async () => {
    helpers.addStyle("Assumption");
    helpers.seed("Data!B2", [["Growth"]]);
    helpers.setStyle("Data!B2", "Assumption");
    helpers.sheet("Data").visibility = "VeryHidden";

    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 2,
      skippedSheets: [],
    });
  });
});

describe("the model check, the corners", () => {
  // A cell holding an error text with no formula behind it is still a cell a
  // reviewer has to look at, and the line carries the error Excel shows.
  it("flags an error value with no formula and names the error", async () => {
    helpers.seed("Model!C7", [["#N/A"]]);

    const { findings } = await smt.runModelCheck();

    expect(of(findings, "formulaError")).toEqual([
      {
        kind: "formulaError",
        sheet: "Model",
        ref: "C7",
        count: 1,
        note: "#N/A",
      },
    ]);
  });

  it("leaves a text cell that only looks like an error alone", async () => {
    helpers.seed("Model!C7", [["not #N/A really"]]);

    expect(of((await smt.runModelCheck()).findings, "formulaError")).toEqual(
      [],
    );
  });
});
