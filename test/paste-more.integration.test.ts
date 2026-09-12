// The three narrow pastes against the fake host: the copied block's formulas
// duplicated to a new corner, its number formats only, its row heights only.
// Each one gets the universality rows - a ctrl-clicked selection, a merged
// block, a protected sheet, a source or selection over the cap, and no copy
// source at all - beside what it is actually for.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
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

// A two by two block whose second column reads its own first column (inside)
// and one cell far away (outside).
async function markBlock(): Promise<void> {
  helpers.seed("Model!A1", [
    [1, { formula: "=A1+10", value: 11 }],
    [2, { formula: "=Z9*B1", value: 0 }],
  ]);
  helpers.select("Model!A1:B2");
  await smt.markCopySource();
}

beforeEach(async () => {
  await boot();
});

describe("paste: duplicate formulas", () => {
  it("adapts in-range references and keeps the external one", async () => {
    await markBlock();
    helpers.select("Model!A5");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Model!B5")).toBe("=A5+10");
    expect(helpers.formula("Model!B6")).toBe("=Z9*B5");
    expect(helpers.value("Model!A5")).toBe(1);
  });

  it("shifts columns as well as rows", async () => {
    await markBlock();
    helpers.select("Model!D5");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Model!E5")).toBe("=D5+10");
    expect(helpers.formula("Model!E6")).toBe("=Z9*E5");
  });

  it("refuses a target that overlaps the copied block", async () => {
    await markBlock();
    helpers.select("Model!B2");

    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Paste target overlaps the copied block.",
    );
    // Nothing was written, so the block still reads as it did.
    expect(helpers.formula("Model!B2")).toBe("=Z9*B1");
  });

  it("allows the same corner on another sheet", async () => {
    await markBlock();
    helpers.select("Data!A1");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Data!B1")).toBe("=A1+10");
    expect(helpers.formula("Data!B2")).toBe("=Z9*B1");
  });

  it("writes into every area of a ctrl-clicked selection", async () => {
    helpers.seed("Model!E1", [[1], [{ formula: "=E1*2", value: 2 }]]);
    helpers.select("Model!E1:E2");
    await smt.markCopySource();

    helpers.selectAreas(["Model!A5", "Model!C5"]);
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Model!A6")).toBe("=A5*2");
    expect(helpers.formula("Model!C6")).toBe("=C5*2");
  });

  it("says what to do when the target cuts a merged cell", async () => {
    await markBlock();
    helpers.merge("Model!B5:C5");
    helpers.select("Model!A5");

    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Paste: Excel refused this write. Select whole merged cells, not part of one.",
    );
  });

  it("names the protected sheet instead of writing", async () => {
    await markBlock();
    helpers.protectSheet("Data");
    helpers.select("Data!A1");

    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Paste: this sheet is protected, nothing was changed",
    );
  });

  it("refuses a source over the cell cap", async () => {
    helpers.select("Model!A1:A5001");
    await smt.markCopySource();
    helpers.select("Data!A1");

    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Paste supports up to 5,000 selected cells at once.",
    );
  });

  it("asks for a source first", async () => {
    helpers.select("Data!A1");
    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Mark a copy source first.",
    );
  });

  it("carries an empty block without complaining", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteDuplicateFormulas();

    expect(helpers.value("Data!B2")).toBe("");
  });

  it("captures the destination for undo", async () => {
    await markBlock();
    helpers.seed("Data!A1", [["keep"]]);
    helpers.select("Data!A1");
    await smt.pasteDuplicateFormulas();
    expect(smt.undoTarget()).toBe("Data!A1:B2");

    await smt.undoLastAction();
    expect(helpers.value("Data!A1")).toBe("keep");
    expect(helpers.formula("Data!B1")).toBe("");
  });
});

describe("paste: number formats only", () => {
  beforeEach(() => {
    helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    helpers.setNumberFormat("Model!A1:B2", "#,##0.0");
    helpers.setFill("Model!A1:B2", { color: "#EEDDCC", pattern: "Solid" });
  });

  it("writes the formats and leaves values and fills alone", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.seed("Data!A1", [[99]]);
    helpers.select("Data!A1");
    await smt.pasteNumberFormats();

    expect(helpers.numberFormat("Data!A1")).toBe("#,##0.0");
    expect(helpers.numberFormat("Data!B2")).toBe("#,##0.0");
    expect(helpers.value("Data!A1")).toBe(99);
    expect(helpers.value("Data!B2")).toBe("");
    expect(helpers.fill("Data!A1").pattern).toBe("None");
  });

  it("tiles a smaller source over the whole selection", async () => {
    helpers.setNumberFormat("Model!A1", "0.0%");
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.select("Data!A1:B2");
    await smt.pasteNumberFormats();

    for (const cell of ["A1", "B1", "A2", "B2"]) {
      expect(helpers.numberFormat(`Data!${cell}`)).toBe("0.0%");
    }
  });

  it("cuts the last tile short on a selection that is not a multiple", async () => {
    helpers.setNumberFormat("Model!A1:A2", "0.00");
    helpers.setNumberFormat("Model!A2", "0.000");
    helpers.select("Model!A1:A2");
    await smt.markCopySource();
    helpers.select("Data!A1:A3");
    await smt.pasteNumberFormats();

    expect(helpers.numberFormat("Data!A1")).toBe("0.00");
    expect(helpers.numberFormat("Data!A2")).toBe("0.000");
    expect(helpers.numberFormat("Data!A3")).toBe("0.00");
  });

  it("grows a one-cell selection to the source shape", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.select("Data!C3");
    await smt.pasteNumberFormats();

    expect(helpers.numberFormat("Data!D4")).toBe("#,##0.0");
  });

  it("writes into every area of a ctrl-clicked selection", async () => {
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.selectAreas(["Data!A1", "Data!C1"]);
    await smt.pasteNumberFormats();

    expect(helpers.numberFormat("Data!A1")).toBe("#,##0.0");
    expect(helpers.numberFormat("Data!C1")).toBe("#,##0.0");
  });

  it("formats a merged block without refusing it", async () => {
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.merge("Data!A1:B1");
    helpers.select("Data!A1:B1");
    await smt.pasteNumberFormats();

    expect(helpers.numberFormat("Data!A1")).toBe("#,##0.0");
  });

  it("names the protected sheet instead of writing", async () => {
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.protectSheet("Data");
    helpers.select("Data!A1");

    expect(await rejects(() => smt.pasteNumberFormats())).toBe(
      "Paste: this sheet is protected, nothing was changed",
    );
  });

  it("refuses a selection over the cell cap", async () => {
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.select("Data!A1:A5001");

    expect(await rejects(() => smt.pasteNumberFormats())).toBe(
      "Paste supports up to 5,000 selected cells at once.",
    );
  });

  it("asks for a source first", async () => {
    helpers.select("Data!A1");
    expect(await rejects(() => smt.pasteNumberFormats())).toBe(
      "Mark a copy source first.",
    );
  });

  it("captures the destination for undo", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.setNumberFormat("Data!A1:B2", "0.000");
    helpers.select("Data!A1");
    await smt.pasteNumberFormats();
    expect(helpers.numberFormat("Data!A1")).toBe("#,##0.0");

    await smt.undoLastAction();
    expect(helpers.numberFormat("Data!A1")).toBe("0.000");
  });
});

describe("paste: row heights only", () => {
  function seedHeights(): void {
    const sheet = helpers.sheet("Model");
    sheet.rowHeights.set(0, 20);
    sheet.rowHeights.set(1, 30);
    sheet.rowHeights.set(2, 40);
    helpers.seed("Model!A1", [[1], [2], [3]]);
  }

  it("writes each source row's height onto the target rows", async () => {
    seedHeights();
    helpers.select("Model!A1:A3");
    await smt.markCopySource();
    helpers.select("Data!B10");

    expect(await smt.pasteRowHeights()).toBe(
      "Paste: 3 row heights (outside pls,fix Undo)",
    );
    expect(helpers.rowHeight("Data", 9)).toBe(20);
    expect(helpers.rowHeight("Data", 10)).toBe(30);
    expect(helpers.rowHeight("Data", 11)).toBe(40);
  });

  it("skips a hidden source row and counts it", async () => {
    seedHeights();
    helpers.hideRows("Model!A2");
    helpers.select("Model!A1:A3");
    await smt.markCopySource();
    helpers.select("Data!A1");

    expect(await smt.pasteRowHeights()).toBe(
      "Paste: 2 row heights, 1 hidden row skipped (outside pls,fix Undo)",
    );
    expect(helpers.rowHeight("Data", 0)).toBe(20);
    // Row 2 of the target keeps Excel's default: the source row was hidden.
    expect(helpers.rowHeight("Data", 1)).toBe(15);
    expect(helpers.rowHeight("Data", 2)).toBe(40);
  });

  it("writes into every area of a ctrl-clicked selection", async () => {
    seedHeights();
    helpers.select("Model!A1:A2");
    await smt.markCopySource();
    helpers.selectAreas(["Data!A1", "Data!C5"]);

    expect(await smt.pasteRowHeights()).toBe(
      "Paste: 2 row heights (outside pls,fix Undo)",
    );
    expect(helpers.rowHeight("Data", 0)).toBe(20);
    expect(helpers.rowHeight("Data", 4)).toBe(20);
    expect(helpers.rowHeight("Data", 5)).toBe(30);
  });

  it("sizes the rows under a merged block", async () => {
    seedHeights();
    helpers.select("Model!A1:A2");
    await smt.markCopySource();
    helpers.merge("Data!A1:B1");
    helpers.select("Data!A1:B1");

    expect(await smt.pasteRowHeights()).toBe(
      "Paste: 2 row heights (outside pls,fix Undo)",
    );
    expect(helpers.rowHeight("Data", 0)).toBe(20);
  });

  it("names the protected sheet instead of writing", async () => {
    seedHeights();
    helpers.select("Model!A1:A2");
    await smt.markCopySource();
    helpers.protectSheet("Data");
    helpers.select("Data!A1");

    expect(await rejects(() => smt.pasteRowHeights())).toBe(
      "Paste: this sheet is protected, nothing was changed",
    );
  });

  it("refuses a source with more rows than one run may carry", async () => {
    helpers.select("Model!A1:A501");
    await smt.markCopySource();
    helpers.select("Data!A1");

    expect(await rejects(() => smt.pasteRowHeights())).toBe(
      "Paste supports up to 500 rows of heights at once.",
    );
  });

  it("asks for a source first", async () => {
    helpers.select("Data!A1");
    expect(await rejects(() => smt.pasteRowHeights())).toBe(
      "Mark a copy source first.",
    );
  });

  it("stays outside pls,fix Undo", async () => {
    seedHeights();
    helpers.select("Model!A1:A2");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteRowHeights();

    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(false);
  });
});
