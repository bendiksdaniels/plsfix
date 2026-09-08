// Slice B audit: the formula, paste and fill flows against the strict fake
// host, on the selections and hosts the green suite never put them on - a
// protected sheet, a merged band, a whole-column click and a selection whose
// active cell is not its first cell. Autocolor and the audit overlay are in
// slice-b.overlay.audit.integration.test.ts.

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

describe("a whole-column click", () => {
  // Excel hands a million cells to anything that asks for their grid, so a flow
  // that reads one says how many it takes first - the way the bridge, the
  // unpivot and the audit overlay already do.
  it("is refused by name before the CAGR reads the values", async () => {
    helpers.select("Model!A:A");
    expect(await rejects(() => smt.insertCagr())).toBe(
      "CAGR supports up to 5,000 selected cells at once.",
    );
  });

  it("is refused by name before an exact paste reads the formulas", async () => {
    helpers.select("Data!A:A");
    await smt.markCopySource();
    helpers.select("Model!B2");

    expect(await rejects(() => smt.pastePreserveFormulas())).toBe(
      "Paste supports up to 5,000 selected cells at once.",
    );
    expect(helpers.formula("Model!B2")).toBe("");
  });
});

describe("the unit and sign edits over a selection a modeller really makes", () => {
  // A label column, numbers, a formula, a percentage, a blank and a flag in one
  // block: only the numbers and the formula may move.
  function seedMixedRow(): void {
    helpers.seed("Model!A1", [
      ["Revenue", 1500, { formula: "=B1*2", value: 3000 }, 0.42, "", true],
    ]);
    helpers.setNumberFormat("Model!D1", "0.0%");
    helpers.select("Model!A1:F1");
  }

  it("scales the numbers and the formula and nothing else", async () => {
    seedMixedRow();
    await smt.scaleSelection(0.001);

    expect(helpers.value("Model!A1")).toBe("Revenue");
    expect(helpers.value("Model!B1")).toBe(1.5);
    expect(helpers.formula("Model!C1")).toBe("=(B1*2)/1000");
    expect(helpers.value("Model!D1")).toBe(0.00042);
    // The percentage keeps its own format; only what it prints changed.
    expect(helpers.numberFormat("Model!D1")).toBe("0.0%");
    expect(helpers.value("Model!E1")).toBe("");
    expect(helpers.value("Model!F1")).toBe(true);
  });

  it("flips the same cells and leaves the label and the flag alone", async () => {
    seedMixedRow();
    await smt.applySignFlip();

    expect(helpers.value("Model!A1")).toBe("Revenue");
    expect(helpers.value("Model!B1")).toBe(-1500);
    expect(helpers.formula("Model!C1")).toBe("=-(B1*2)");
    expect(helpers.value("Model!F1")).toBe(true);
  });

  it("leaves a cell holding an error value where it is", async () => {
    helpers.seed("Model!A1", [
      ["#REF!", { formula: "=1/0", value: "#DIV/0!" }, "#N/A"],
    ]);
    helpers.select("Model!A1:C1");

    await smt.applySignFlip();
    expect(helpers.formula("Model!A1")).toBe("#REF!");
    expect(helpers.formula("Model!C1")).toBe("#N/A");

    await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A1")).toBe("#REF!");
    expect(helpers.formula("Model!B1")).toBe("=IFERROR(-(1/0),0)");
  });
});

describe("a rounding group past its cap", () => {
  it("says how many cells it takes", async () => {
    helpers.select("Model!A1:A2000");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding groups up to 1000 cells at once.",
    );
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
