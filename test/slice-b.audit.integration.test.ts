// Slice B audit: the formula, paste, fill, autocolor, audit-overlay and trace
// flows against the strict fake host, on the selections and hosts the green
// suite never put them on - a protected sheet, a merged band, a selection whose
// active cell is not its first cell, a corrupt overlay snapshot in the file.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
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

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------

describe("an edit on a protected sheet", () => {
  // Every flow that writes says which stage was refused and that nothing
  // changed; office.js's own "The worksheet Model is protected." never reaches
  // the pane, the way protection.ts's header promises.
  async function markSource(): Promise<void> {
    helpers.seed("Data!A1", [[1, { formula: "=A1+1", value: 2 }]]);
    helpers.select("Data!A1:B1");
    await smt.markCopySource();
  }

  it("names the paste that was refused", async () => {
    await markSource();
    helpers.select("Model!A1:B1");
    helpers.protectSheet("Model");

    for (const mode of ["values", "formats", "transpose"] as const) {
      expect(await rejects(() => smt.pasteSpecial(mode))).toBe(
        "Paste: this sheet is protected, nothing was changed",
      );
    }
    expect(await rejects(() => smt.pastePreserveFormulas())).toBe(
      "Paste: this sheet is protected, nothing was changed",
    );
    expect(helpers.value("Model!A1")).toBe("");
  });

  it("names the fill that was refused", async () => {
    helpers.seed("Model!B1", [["Q1", "Q2", "Q3"]]);
    helpers.seed("Model!B2", [[{ formula: "=A2*2", value: 8 }]]);
    helpers.select("Model!B2");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.fastFillAuto("right"))).toBe(
      "Fill: this sheet is protected, nothing was changed",
    );
    expect(helpers.formula("Model!C2")).toBe("");
  });

  it("names the CAGR that was refused", async () => {
    helpers.seed("Model!A2", [[100, 110, 121]]);
    helpers.select("Model!A2:C2");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.insertCagr())).toBe(
      "CAGR: this sheet is protected, nothing was changed",
    );
    expect(helpers.formula("Model!D2")).toBe("");
  });

  it("names the rounding block that was refused", async () => {
    helpers.seed("Model!B2", [[33.3, 33.3, 33.4]]);
    helpers.select("Model!B2:D2");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding: this sheet is protected, nothing was changed",
    );
    expect(helpers.formula("Model!B3")).toBe("");
  });
});

describe("a paste that cuts a merged cell", () => {
  it("says what to do instead of handing back Excel's string", async () => {
    helpers.seed("Data!A1", [[{ formula: "=Z1+1", value: 2 }], [7]]);
    helpers.select("Data!A1:A2");
    await smt.markCopySource();

    helpers.seed("Model!A1", [["Revenue bridge"]]);
    helpers.merge("Model!A1:B1");
    // Column A only: the merged A1:B1 band is half in, half out.
    helpers.select("Model!A1");

    expect(await rejects(() => smt.pastePreserveFormulas())).toBe(
      "Paste: Excel refused this write. Select whole merged cells, not part of one.",
    );
    expect(helpers.value("Model!A1")).toBe("Revenue bridge");
  });
});

// ---------------------------------------------------------------------------

describe("a fast fill whose active cell is not the selection's first cell", () => {
  // Dragging a selection upwards or leftwards leaves Excel's active cell at the
  // far corner. The fill starts there, so the selection may only size it as far
  // as its own edge - never past it, and never over cells nobody selected.
  it("fills down to the end of the selection, not past it", async () => {
    helpers.seed("Model!B3", [[{ formula: "=A3*2", value: 2 }]]);
    helpers.select("Model!B1:B5");
    helpers.setActiveCell("Model!B3");

    await smt.fastFillAuto("down");

    expect(helpers.formula("Model!B5")).toBe("=A3*2");
    expect(helpers.formula("Model!B6")).toBe("");
    expect(helpers.formula("Model!B7")).toBe("");
  });

  it("fills right to the end of the selection, not past it", async () => {
    helpers.seed("Model!C1", [[{ formula: "=C9*2", value: 2 }]]);
    helpers.select("Model!A1:D1");
    helpers.setActiveCell("Model!C1");

    await smt.fastFillAuto("right");

    expect(helpers.formula("Model!D1")).toBe("=C9*2");
    expect(helpers.formula("Model!E1")).toBe("");
  });

  it("refuses when the active cell sits on the selection's last row", async () => {
    helpers.seed("Model!B5", [[{ formula: "=A5*2", value: 2 }]]);
    helpers.select("Model!B1:B5");
    helpers.setActiveCell("Model!B5");

    expect(await rejects(() => smt.fastFillAuto("down"))).toBe(
      "No neighbor data to size the fill.",
    );
    expect(helpers.formula("Model!B6")).toBe("");
  });

  it("still lets the neighbour data reach past the selection", async () => {
    helpers.seed("Model!B1", [["Q1", "Q2", "Q3", "Q4"]]);
    helpers.seed("Model!B2", [[{ formula: "=A2*2", value: 8 }]]);
    helpers.select("Model!B2:C2");

    await smt.fastFillAuto("right");

    expect(helpers.formula("Model!E2")).toBe("=A2*2");
  });
});
