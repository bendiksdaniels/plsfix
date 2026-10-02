// Pass-1 hunt: where a template block may land and what stops it: the sheet
// edge, the active sheet, a quoted sheet name, a formula showing nothing, a
// comment or formatting only, a protected sheet, three presses then Undo.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
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
// The sheet edge: a block must refuse cleanly, never write the part of itself
// that happened to fit.
// ---------------------------------------------------------------------------

describe("a block that would run past the sheet edge", () => {
  it("refuses past the last row, and writes nothing at all", async () => {
    // debt-schedule is 20 rows; anchored 5 rows from the bottom it needs 15
    // more than the sheet has left.
    const anchor = cellAddress(1_048_576 - 5, 0);
    helpers.select(`Model!${anchor}`);

    expect(await rejects(() => smt.insertTemplate("debt-schedule"))).toBe(
      "Templates need an empty block of 20 rows by 6 columns, and there is no room for it at the selection.",
    );
    expect(helpers.value(`Model!${anchor}`)).toBe("");
  });

  it("refuses past the last column (XFD), and writes nothing at all", async () => {
    // debt-schedule is 6 columns wide; anchored 4 columns from the right edge
    // it needs 2 more than the sheet has left.
    const anchor = cellAddress(0, 16_384 - 4);
    helpers.select(`Model!${anchor}`);

    expect(await rejects(() => smt.insertTemplate("debt-schedule"))).toBe(
      "Templates need an empty block of 20 rows by 6 columns, and there is no room for it at the selection.",
    );
    expect(helpers.value(`Model!${anchor}`)).toBe("");
  });

  it("still writes when the block fits exactly against the edge", async () => {
    // ebitda-bridge is 9x2: anchored so it ends exactly on the last row.
    const anchor = cellAddress(1_048_576 - 9, 16_384 - 2);
    helpers.select(`Model!${anchor}`);

    expect(await smt.insertTemplate("ebitda-bridge")).toBe(
      "Template written: EBITDA bridge (9x2)",
    );
    expect(helpers.value(`Model!${anchor}`)).toBe("EBITDA bridge");
  });
});

// ---------------------------------------------------------------------------
// Debt schedule: PMT sized off the inputs must fully amortise the principal.
// ---------------------------------------------------------------------------

describe("placement quirks", () => {
  it("lands on whichever sheet is active, not always the first one", async () => {
    await boot(["Model", "Data"]);
    helpers.select("Data!B2");

    expect(await smt.insertTemplate("working-capital")).toBe(
      "Template written: Working capital days (12x2)",
    );
    expect(helpers.value("Data!B2")).toBe("Working capital days");
    // The other sheet, same address, is untouched.
    expect(helpers.value("Model!B2")).toBe("");
  });

  it("writes formulas with no sheet prefix at all on a sheet whose name needs quoting", async () => {
    await boot(["Model", "Q1 Plan's Data"]);
    helpers.select("Q1 Plan's Data!A1");

    expect(await smt.insertTemplate("ebitda-bridge")).toBe(
      "Template written: EBITDA bridge (9x2)",
    );
    // Every placeholder in every template is an intra-block reference, so
    // nothing here should ever carry a sheet prefix (quoted or not).
    expect(helpers.formula("Q1 Plan's Data!B9")).toBe("=B2+SUM(B3:B7)-B8");
  });
});

// ---------------------------------------------------------------------------
// Block occupancy: what actually blocks a template, and what does not.
// ---------------------------------------------------------------------------

describe("what blocks an otherwise-empty target block", () => {
  it("refuses over a formula that shows nothing, and leaves it untouched", async () => {
    helpers.seed("Model!B5", [[{ formula: '=IF(1=2,"","")', value: "" }]]);
    helpers.select("Model!A1");

    expect(await rejects(() => smt.insertTemplate("ebitda-bridge"))).toBe(
      "Templates need an empty block of 9 rows by 2 columns at the selection.",
    );
    expect(helpers.formula("Model!B5")).toBe('=IF(1=2,"","")');
  });

  it("does not refuse over a cell that carries only a comment", async () => {
    helpers.addComment("Model!B5", "check this line", "Reviewer");
    helpers.select("Model!A1");

    expect(await smt.insertTemplate("ebitda-bridge")).toBe(
      "Template written: EBITDA bridge (9x2)",
    );
  });

  it("does not refuse over a cell that carries only formatting", async () => {
    helpers.setFill("Model!B5", { color: "#FFFF00" });
    helpers.select("Model!A1");

    expect(await smt.insertTemplate("ebitda-bridge")).toBe(
      "Template written: EBITDA bridge (9x2)",
    );
  });
});

// ---------------------------------------------------------------------------
// Protected sheet.
// ---------------------------------------------------------------------------

describe("a protected sheet", () => {
  it("names the sheet and changes nothing, instead of Excel's own AccessDenied", async () => {
    helpers.protectSheet("Model");
    helpers.select("Model!A1");

    expect(await rejects(() => smt.insertTemplate("ebitda-bridge"))).toBe(
      "Templates: this sheet is protected, nothing was changed",
    );
    expect(helpers.value("Model!A1")).toBe("");
  });
});

// ---------------------------------------------------------------------------
// Pressed three times, then Undo.
// ---------------------------------------------------------------------------

describe("pressed three times in a row", () => {
  it("writes once, refuses twice more over its own block, then undo restores exactly", async () => {
    helpers.select("Model!A1");

    expect(await smt.insertTemplate("ebitda-bridge")).toBe(
      "Template written: EBITDA bridge (9x2)",
    );
    const afterFirst = helpers.cellMap("Model");

    const refusal =
      "Templates need an empty block of 9 rows by 2 columns at the selection.";
    expect(await rejects(() => smt.insertTemplate("ebitda-bridge"))).toBe(
      refusal,
    );
    expect(helpers.cellMap("Model")).toEqual(afterFirst);

    expect(await rejects(() => smt.insertTemplate("ebitda-bridge"))).toBe(
      refusal,
    );
    expect(helpers.cellMap("Model")).toEqual(afterFirst);

    expect(await smt.undoLastAction()).toBe(
      "Undone: Model!A1:B9. Nothing more to undo.",
    );
    expect(helpers.value("Model!A1")).toBe("");
  });
});
