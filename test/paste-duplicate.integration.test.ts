// The duplicate-formula paste against the fake host: the copied block's
// formulas at a new corner, in-range references adapted and outside ones still
// reading the cells they were written for. Universality rows included: a
// ctrl-clicked selection, a merged block, a protected sheet, a source over the
// cap and no copy source at all.

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

  it("qualifies the outside reference when the paste lands on another sheet", async () => {
    await markBlock();
    helpers.select("Data!A1");
    await smt.pasteDuplicateFormulas();

    // In-block: the destination's own cells, so still unqualified.
    expect(helpers.formula("Data!B1")).toBe("=A1+10");
    // Outside: Model!Z9 is the cell it was written for, and saying so is the
    // only way it keeps pointing there from a sheet that is not Model.
    expect(helpers.formula("Data!B2")).toBe("=Model!Z9*B1");
  });

  it("re-points an in-block reference written with the source sheet's name", async () => {
    helpers.seed("Model!A1", [
      [1, { formula: "=Model!A1+Model!Z9", value: 1 }],
    ]);
    helpers.select("Model!A1:B1");
    await smt.markCopySource();
    helpers.select("Data!A5");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Data!B5")).toBe("=Data!A5+Model!Z9");
  });

  it("quotes a source sheet name that needs it", async () => {
    helpers.addSheet("P&L 2025");
    helpers.seed("P&L 2025!A1", [[{ formula: "=Z9", value: 0 }]]);
    helpers.select("P&L 2025!A1");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Data!A1")).toBe("='P&L 2025'!Z9");
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
