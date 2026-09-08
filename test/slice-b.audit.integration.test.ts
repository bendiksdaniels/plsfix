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
import type * as SettingsModule from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;
let theme: ReturnType<(typeof SettingsModule)["deriveTheme"]>;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
  const brand = await import("../src/settings");
  theme = brand.deriveTheme(brand.DEFAULT_SETTINGS);
}

// Reopening the file: same workbook model, brand new runtime and module state.
async function reopen(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ workbook });
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

// ---------------------------------------------------------------------------

describe("autocolor over a structured table reference", () => {
  // =SUM(Table1[Revenue]) points at a table in this workbook, not at another
  // file: painting it as an external link tells a reviewer the model reaches
  // outside when it does not.
  it("paints a table reference as a formula, not an external link", async () => {
    helpers.seed("Model!A1", [
      [{ formula: "=SUM(Table1[Revenue])", value: 10 }],
      [{ formula: "=SUM(Sales[[#Headers],[Amount]])", value: 4 }],
      [{ formula: "=[Model.xlsx]Sheet1!A1", value: 1 }],
    ]);
    helpers.select("Model!A1:A3");

    expect(await smt.autocolorSelection()).toBe("Autocolor: 3 cells");
    expect(helpers.font("Model!A1").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!A2").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!A3").color).toBe(theme.externalFont);
  });

  it("counts one cell in the singular", async () => {
    helpers.seed("Model!A1", [[1234]]);
    helpers.select("Model!A1");
    expect(await smt.autocolorSelection()).toBe("Autocolor: 1 cell");
  });
});

describe("autocolor on edit", () => {
  it("switching off when it was never on touches no handler", async () => {
    await smt.setAutocolorOnEdit(false);
    expect(helpers.changeHandlerCount()).toBe(0);
    expect(smt.autocolorOnEditActive()).toBe(false);
  });

  it("does not recolour its own paint while one is in flight", async () => {
    helpers.seed("Data!A1", [[42]]);
    await smt.setAutocolorOnEdit(true);

    // Both events arrive before the first has finished its round trips; the
    // second one finds the flag set and returns without a second Excel.run.
    await Promise.all([
      helpers.fireChanged("Data", "A1"),
      helpers.fireChanged("Data", "A1"),
    ]);
    expect(helpers.font("Data!A1").color).toBe(theme.inputFont);
  });
});

// ---------------------------------------------------------------------------

describe("a corrupt audit overlay snapshot in the file", () => {
  // The boot restore is the only thing that can take last session's stripes
  // off, and src/main.ts swallows whatever it throws: a snapshot it cannot read
  // has to be dropped, or the model stays striped on every reopen.
  it("drops a snapshot that is not JSON at all", async () => {
    helpers.setSetting("smtAuditOverlay", "not json at all");

    expect(await smt.restorePersistedOverlay()).toBe(false);
    expect(helpers.setting("smtAuditOverlay")).toBe("");
  });

  it("drops a snapshot that is JSON but not a list", async () => {
    helpers.setSetting("smtAuditOverlay", '{"address":"A1:C3"}');

    expect(await smt.restorePersistedOverlay()).toBe(false);
    expect(helpers.setting("smtAuditOverlay")).toBe("");
  });
});

// ---------------------------------------------------------------------------

// The pane reloads mid-overlay: the map dies with the runtime while the stripes
// were saved with the file, so the copy in workbook.settings is the only way
// back to the modeller's own fills. src/main.ts calls restorePersistedOverlay
// at boot; this is the store side of that, end to end.
describe("the audit overlay across a pane reload", () => {
  function seedBlock(): void {
    helpers.seed("Model!A1", [
      [
        { formula: "=B1*2", r1c1: "=RC[1]*2", value: 2 },
        { formula: "=C1*2", r1c1: "=RC[1]*2", value: 4 },
      ],
      [
        { formula: "=B2*2", r1c1: "=RC[1]*2", value: 6 },
        { formula: "=X9", r1c1: "=X9", value: 1 },
      ],
    ]);
  }

  it("puts the modeller's own fills back and forgets the snapshot", async () => {
    seedBlock();
    helpers.setFill("Model!A1", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    const before = helpers.cellMap("Model");

    helpers.select("Model!A1:B2");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    expect(smt.auditOverlayOn()).toBe(true);
    expect(helpers.fill("Model!A1").pattern).not.toBe("LightUp");

    await reopen();
    expect(smt.auditOverlayOn()).toBe(false);
    expect(await smt.restorePersistedOverlay()).toBe(true);
    expect(helpers.cellMap("Model")).toEqual(before);
    expect(helpers.setting("smtAuditOverlay")).toBe("");

    // Nothing is owned any more, so the next toggle is a first toggle: it
    // snapshots the modeller's fills, not the stripes it put there before.
    helpers.select("Model!A1:B2");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("has nothing left to restore on the boot after that", async () => {
    seedBlock();
    helpers.select("Model!A1:B2");
    await smt.toggleAuditOverlay();

    await reopen();
    expect(await smt.restorePersistedOverlay()).toBe(true);
    await reopen();
    expect(await smt.restorePersistedOverlay()).toBe(false);
  });

  it("drops a snapshot whose sheet was deleted in between", async () => {
    seedBlock();
    helpers.select("Model!A1:B2");
    await smt.toggleAuditOverlay();

    helpers.deleteSheet("Model");
    await reopen();
    expect(await smt.restorePersistedOverlay()).toBe(false);
    expect(helpers.setting("smtAuditOverlay")).toBe("");
    expect(workbook.sheets.length).toBe(1);
  });
});
