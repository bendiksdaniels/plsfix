// Pass-1 hunt: sheet names with apostrophes crossing through
// src/formula-duplicate.ts and the full src/excel/paste.ts adapter, plus a
// whole-row/whole-column selection meeting the paste cap where the code
// actually enforces it (the source's shape, not the raw target selection).
import { beforeEach, describe, expect, it, vi } from "vitest";
import { duplicateFormula } from "../../src/formula-duplicate";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

// A1:B1 on a sheet whose name needs quoting and apostrophe-doubling.
const BLOCK = {
  sheet: "O'Brien",
  row: 0,
  column: 0,
  rowCount: 1,
  columnCount: 2,
};
const RIGHT = { rows: 0, columns: 1 };
// The one way Excel actually writes this sheet name into a formula: quoted,
// the apostrophe doubled. A bare "O'Brien!" (no quotes) is not valid syntax
// and the scanner reads no prefix out of it at all.
const QUOTED = "'O''Brien'!";

describe("duplicateFormula: a sheet name with an apostrophe", () => {
  it("keeps an outside reference on its own sheet when pasting onto O'Brien", () => {
    // Pasting FROM "Data" ONTO "O'Brien": the outside reference has to name
    // "Data" to keep pointing there, unaffected by the destination's name.
    const block = {
      sheet: "Data",
      row: 0,
      column: 0,
      rowCount: 1,
      columnCount: 1,
    };
    expect(
      duplicateFormula("=Z9", block, { rows: 0, columns: 1 }, "O'Brien"),
    ).toBe("=Data!Z9");
  });

  it("keeps an in-block self-qualified reference qualified on a same-sheet paste", () => {
    // Same rule as a plain name (formula-duplicate.test.ts "=Model!B2" ->
    // "=Model!B6"): an explicit same-sheet prefix survives, it is not
    // stripped to bare "=B1".
    expect(duplicateFormula(`=${QUOTED}A1`, BLOCK, RIGHT, "O'Brien")).toBe(
      `=${QUOTED}B1`,
    );
  });

  it("re-points an in-block reference written with the apostrophe'd source name", () => {
    // Crossing onto a different sheet, the same in-block reference becomes
    // the destination's own cell, written with the destination's name.
    expect(duplicateFormula(`=${QUOTED}A1`, BLOCK, RIGHT, "Data")).toBe(
      "=Data!B1",
    );
  });

  it("qualifies an outside reference with the quoted, doubled source name", () => {
    expect(duplicateFormula("=Z9", BLOCK, RIGHT, "Data")).toBe(`=${QUOTED}Z9`);
  });
});

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
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

describe("pasteDuplicateFormulas end to end onto an apostrophe'd sheet", () => {
  it("writes the properly quoted outside reference through the real adapter", async () => {
    helpers.addSheet("O'Brien");
    helpers.seed("Model!A1", [[1, { formula: "=Z9", value: 0 }]]);
    helpers.select("Model!A1:B1");
    await smt.markCopySource();

    helpers.select("O'Brien!A1");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("O'Brien!B1")).toBe("=Model!Z9");
  });

  it("re-points a reference already qualified with the apostrophe'd source, keeps the outside one", async () => {
    helpers.addSheet("O'Brien");
    helpers.seed("O'Brien!A1", [
      [1, { formula: `=${QUOTED}A1+${QUOTED}Z9`, value: 1 }],
    ]);
    helpers.select("O'Brien!A1:B1");
    await smt.markCopySource();

    helpers.select("Model!A5");
    await smt.pasteDuplicateFormulas();

    expect(helpers.formula("Model!B5")).toBe(`=Model!A5+${QUOTED}Z9`);
  });
});

describe("a whole row or column meets the paste cap where the code enforces it", () => {
  it("marks a whole-column source without complaint, then refuses it at paste time", async () => {
    // markCopySource records the address only; the cap is the source's
    // cellCount, read fresh inside the paste call that actually reads it.
    helpers.select("Model!A:A");
    await expect(smt.markCopySource()).resolves.toBe("Model!A:A");

    helpers.select("Model!C1");
    expect(await rejects(() => smt.pasteDuplicateFormulas())).toBe(
      "Paste supports up to 5,000 selected cells at once.",
    );
  });

  it("writes only the source's own shape at a whole row's start, not the whole row", async () => {
    // pasteDuplicateFormulas sizes the write from the SOURCE, not the raw
    // target selection, so a one-cell source onto a whole-row destination
    // writes one cell at the row's start rather than 16,384 of them.
    helpers.seed("Model!A1", [[{ formula: "=Z9", value: 0 }]]);
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.select("Model!5:5");

    await smt.pasteDuplicateFormulas();
    expect(helpers.formula("Model!A5")).toBe("=Z9");
    expect(helpers.formula("Model!B5")).toBe("");
  });

  it("refuses a whole-column number-format destination cleanly (that one IS capped by selection)", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    await smt.markCopySource();
    helpers.select("Model!C:C");

    expect(await rejects(() => smt.pasteNumberFormats())).toBe(
      "Paste supports up to 5,000 selected cells at once.",
    );
  });
});

describe("the same paste action pressed onto two different sheets in a row", () => {
  it("does not leak state between an unrelated sheet and the real target", async () => {
    helpers.addSheet("Extra");
    helpers.seed("Model!A1", [[1, { formula: "=A1*2", value: 2 }]]);
    helpers.select("Model!A1:B1");
    await smt.markCopySource();

    helpers.select("Extra!A1");
    await smt.pasteDuplicateFormulas();
    expect(helpers.formula("Extra!B1")).toBe("=A1*2");

    // The copy source is remembered by sheet id and address, so pasting a
    // second time - same source, a different sheet - must read the same
    // block again, not whatever the first paste just wrote.
    helpers.select("Model!D1");
    await smt.pasteDuplicateFormulas();
    expect(helpers.formula("Model!E1")).toBe("=D1*2");
  });
});
