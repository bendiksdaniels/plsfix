// Slice B audit, second half: autocolor and the audit overlay against the
// strict fake host - what a structured table reference is painted as, what the
// overlay may own, and what a snapshot saved in the file does on the next boot.
// The formula, paste and fill flows are in slice-b.audit.integration.test.ts.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
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

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
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

  it("refuses a whole column by its cell count, on any host", async () => {
    helpers.select("Model!A:A");
    expect(await rejects(() => smt.autocolorSelection())).toBe(
      "Autocolor supports up to 5,000 selected cells at once.",
    );

    // Below ExcelApi 1.9 there are no RangeAreas, so the count comes off the
    // one selected block instead - and still before any grid is asked for.
    await boot();
    helpers.setSupported((_set, version) => version !== "1.9");
    helpers.select("Model!A:A");
    expect(await rejects(() => smt.autocolorSelection())).toBe(
      "Autocolor supports up to 5,000 selected cells at once.",
    );
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

describe("the audit overlay over a block with no formulas", () => {
  // An overlay that owns fills it never painted reads as "On" with nothing on
  // the sheet, saves a snapshot in the file and locks the linked-cell highlight
  // out until it is toggled off again.
  it("stays off and says why on a blank cell", async () => {
    helpers.select("Model!F9");

    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(smt.lastAuditNote()).toBe("no formula here to stripe");
    expect(smt.auditOverlayOn()).toBe(false);
    expect(helpers.setting("smtAuditOverlay")).toBeNull();
  });

  it("stays off over a block of typed numbers", async () => {
    helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    const before = helpers.cellMap("Model");
    helpers.select("Model!A1:B2");

    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(smt.auditOverlayOn()).toBe(false);
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("still takes the fills back off a block it did stripe", async () => {
    helpers.seed("Model!A1", [
      [
        { formula: "=B1*2", r1c1: "=RC[1]*2", value: 2 },
        { formula: "=C1*2", r1c1: "=RC[1]*2", value: 4 },
      ],
    ]);
    const before = helpers.cellMap("Model");
    helpers.select("Model!A1:B1");

    expect(await smt.toggleAuditOverlay()).toBe(true);
    expect(smt.lastAuditNote()).toBeNull();
    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(helpers.cellMap("Model")).toEqual(before);
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

  it("keeps only the well-formed entries of a list", async () => {
    helpers.setSetting(
      "smtAuditOverlay",
      '[1, {"address":"A1:C3"}, {"sheetId":"nowhere","address":"A1","cells":[]}]',
    );

    // A number and a half-written object are dropped before any range is
    // asked for; the one snapshot with all three parts is restored or, its
    // sheet gone, discarded like any stale entry. Nothing throws.
    await expect(smt.restorePersistedOverlay()).resolves.not.toThrow();
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
