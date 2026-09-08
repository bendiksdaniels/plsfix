// Audit of the selection formatters against the strict fake host: what a
// protected sheet says, what a refused row-style cycle does to pls,fix Undo,
// the pre-1.9 single-area fallback, and the promise that formatting never
// depends on local storage or on the Links tab having been opened.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  hostError,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let brand: typeof SettingsModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
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

describe("a refused row-style cycle", () => {
  // The cap is a refusal, not an action: it must leave the five-deep stack
  // exactly as the last real action left it.
  it("keeps the undo slot the previous action filled", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    await smt.applyPreset("input");
    expect(smt.undoTarget()).toBe("Model!A1");

    helpers.select("Model!A1:A600");
    expect(await rejects(() => smt.applyRowStyleCycle("item"))).toBe(
      "Row styles support up to 500 rows at once.",
    );

    expect(smt.undoTarget()).toBe("Model!A1");
    expect(smt.lastUndoSkipped()).toBe(false);
  });

  // A whole-column click is over the capture's own cell cap, which drops the
  // stack and arms the "too large for undo" note - for an action that wrote
  // nothing at all.
  it("keeps the stack when the selection is a whole column", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    await smt.applyPreset("input");

    helpers.select("Model!A:A");
    expect(await rejects(() => smt.applyRowStyleCycle("title"))).toBe(
      "Row styles support up to 500 rows at once.",
    );

    expect(smt.undoTarget()).toBe("Model!A1");
    expect(smt.lastUndoSkipped()).toBe(false);
    await smt.undoLastAction();
    expect(helpers.font("Model!A1").color).toBe("#000000");
  });

  it("still cycles a selection inside the cap", async () => {
    helpers.select("Model!A1:A500");
    await smt.applyRowStyleCycle("item");
    expect(helpers.font("Model!A500").color).toBe(
      brand.deriveTheme(brand.DEFAULT_SETTINGS).formulaFont,
    );
  });
});

// ---------------------------------------------------------------------------

describe("formatting a protected sheet", () => {
  beforeEach(() => {
    helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!A1:B2");
    helpers.protectSheet("Model");
  });

  // Every one of these used to travel as office.js's own "The worksheet Model
  // is protected." - no stage, nothing the modeller can act on.
  const refusals: [string, string][] = [
    ["Formatting", "preset"],
    ["Clearing formats", "eraser"],
    ["Fill cycling", "fill cycle"],
    ["Font colour cycling", "font cycle"],
    ["Border cycling", "border cycle"],
    ["Row styles", "row style cycle"],
    ["Format cycling", "number cycle"],
  ];

  function run(what: string): Promise<unknown> {
    switch (what) {
      case "preset":
        return smt.applyPreset("title");
      case "eraser":
        return smt.clearFormats();
      case "fill cycle":
        return smt.applyFillCycle();
      case "font cycle":
        return smt.applyFontColorCycle();
      case "border cycle":
        return smt.applyBorderCycle();
      case "row style cycle":
        return smt.applyRowStyleCycle("title");
      default:
        return smt.applyNumberCycle("general");
    }
  }

  for (const [stage, what] of refusals) {
    it(`names itself when the ${what} is refused`, async () => {
      const before = helpers.cellMap("Model");
      expect(await rejects(() => run(what))).toBe(
        `${stage}: this sheet is protected, nothing was changed`,
      );
      expect(helpers.cellMap("Model")).toEqual(before);
    });
  }
});

describe("undo into a sheet that was protected afterwards", () => {
  it("names itself instead of handing back Excel's string", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    await smt.applyPreset("input");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.undoLastAction())).toBe(
      "Undo: this sheet is protected, nothing was changed",
    );
    // A failed restore stays retryable, so the entry is still there.
    expect(smt.undoTarget()).toBe("Model!A1");
  });
});

// ---------------------------------------------------------------------------

// Ctrl+A. Range.cellCount answers -1 once the count runs past 2^31-1, and a
// whole sheet is 17.2e9 cells, so every cap that only asked "> cap" read the
// biggest possible selection as "no cells at all" and went on to read it.
describe("a whole-sheet selection", () => {
  const WHOLE_SHEET = "Model!A1:XFD1048576";
  const CAP = "supports up to 5,000 selected cells at once.";

  beforeEach(() => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select(WHOLE_SHEET);
  });

  it("is over the cap for a single range", async () => {
    const { withinCap } = await import("../src/excel/internal");

    await expect(
      Excel.run((context) =>
        withinCap(context, context.workbook.getSelectedRange(), "Paintbrush"),
      ),
    ).rejects.toThrow(`Paintbrush ${CAP}`);
  });

  it("is over the cap for the areas of a selection", async () => {
    const { cappedAreas } = await import("../src/excel/areas");

    await expect(
      Excel.run((context) => cappedAreas(context, "Border cycling")),
    ).rejects.toThrow(`Border cycling ${CAP}`);
  });

  // Every one of these used to reach makeFormatGrid with the whole grid's
  // dimensions and take the pane's webview down with it.
  it("is refused by the flows that write a grid", async () => {
    expect(await rejects(() => smt.applyNumberFormat("whole"))).toBe(
      `Number formatting ${CAP}`,
    );
    expect(await rejects(() => smt.applyNumberCycle("general"))).toBe(
      `Format cycling ${CAP}`,
    );

    helpers.setActiveCell("Model!A1");
    const slot = await smt.captureSlot(1);
    helpers.select(WHOLE_SHEET);
    expect(await rejects(() => smt.applySlot(1, slot))).toBe(
      `Paintbrush ${CAP}`,
    );
  });

  // The selection card is a passive read after every action, so it must not be
  // the thing that asks the host for seventeen billion cells.
  it("is shown by address and count only", async () => {
    const summary = await smt.inspectSelection();

    expect(summary.address).toContain("A1:XFD1048576");
    expect(summary.formulas).toBe(-1);
    expect(summary.blanks).toBe(-1);
  });
});

// ---------------------------------------------------------------------------

// protection.ts is the one place a refusal is translated. Everything else has
// to travel untouched, or a broken host would read as a locked sheet.
describe("the write guard itself", () => {
  it("passes an unrelated failure through both wrappers", async () => {
    const { paintSync, syncWrite } = await import("../src/excel/protection");

    helpers.failNextSync(hostError("GeneralException", "Something broke."));
    await expect(
      Excel.run((context) => syncWrite(context, "Formatting")),
    ).rejects.toThrow("Something broke.");

    helpers.failNextSync(hostError("GeneralException", "Something broke."));
    await expect(
      Excel.run((context) => paintSync(context, "Autocolor", "done")),
    ).rejects.toThrow("Something broke.");
  });

  it("reports a paint refusal as the note and a clean batch as done", async () => {
    const { paintSync } = await import("../src/excel/protection");

    helpers.failNextSync(
      hostError("AccessDenied", "The worksheet Model is protected."),
    );
    await expect(
      Excel.run((context) => paintSync(context, "Autocolor", "done")),
    ).resolves.toBe("Autocolor: this sheet is protected, nothing was changed");

    await expect(
      Excel.run((context) => paintSync(context, "Autocolor", "done")),
    ).resolves.toBe("done");
  });

  // The waterfall's surface batch: Excel for the web rejects it whole, and the
  // chart has to survive that one refusal without swallowing anything else.
  it("tolerates one named code and nothing else", async () => {
    const { syncTolerating } = await import("../src/excel/internal");

    helpers.failNextSync(hostError("UnsupportedOperation", "not here"));
    await expect(
      Excel.run((context) => syncTolerating(context, "UnsupportedOperation")),
    ).resolves.toBe(false);

    await expect(
      Excel.run((context) => syncTolerating(context, "UnsupportedOperation")),
    ).resolves.toBe(true);

    helpers.failNextSync(hostError("GeneralException", "Something broke."));
    await expect(
      Excel.run((context) => syncTolerating(context, "UnsupportedOperation")),
    ).rejects.toThrow("Something broke.");
  });

  // Excel.WorksheetProtection is ExcelApi 1.2. A host too old to be asked
  // answers false, and the write itself is what finds out.
  it("answers false on a host that cannot be asked", async () => {
    await boot({
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.2"),
    });
    const { sheetProtected } = await import("../src/excel/protection");
    helpers.protectSheet("Model");

    const asked = await Excel.run((context) =>
      sheetProtected(context, context.workbook.worksheets.getItem("Model")),
    );
    expect(asked).toBe(false);
  });
});

// ---------------------------------------------------------------------------

// getSelectedRanges is ExcelApi 1.9. Below it the adapter serves the single
// rectangle getSelectedRange has always answered, and turns the host's bare
// InvalidSelection into a sentence naming the flow that asked.
describe("a host below ExcelApi 1.9", () => {
  const old = (set: string, version: string): boolean =>
    set !== "ExcelApi" || Number(version) <= 1.8;

  beforeEach(async () => {
    await boot({ isSetSupported: old });
  });

  it("formats the one block it can serve", async () => {
    helpers.seed("Model!A1", [[1], [2]]);
    helpers.select("Model!A1:A2");
    await smt.applyPreset("input");

    expect(helpers.font("Model!A2").color).toBe(
      brand.deriveTheme(brand.DEFAULT_SETTINGS).inputFont,
    );
    expect(smt.undoTarget()).toBe("Model!A1:A2");
  });

  it("says which flow could not take a ctrl-clicked selection", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.selectAreas(["Model!A1:A2", "Model!C1:C2"]);

    expect(await rejects(() => smt.applyPreset("input"))).toBe(
      "Formatting: this Excel build can only act on one selected block.",
    );
    expect(await rejects(() => smt.applyNumberFormat("whole"))).toBe(
      "Number formatting: this Excel build can only act on one selected block.",
    );
    expect(await rejects(() => smt.applyBorderCycle())).toBe(
      "Border cycling: this Excel build can only act on one selected block.",
    );
  });

  it("still refuses a selection over the cell cap", async () => {
    helpers.select("Model!A:A");
    expect(await rejects(() => smt.applyNumberFormat("whole"))).toBe(
      "Number formatting supports up to 5,000 selected cells at once.",
    );
  });

  // Only InvalidSelection means "several blocks"; anything else is a host
  // failure and has to keep its own message.
  it("passes an unrelated host failure through untouched", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    helpers.failNextSync(hostError("GeneralException", "Something broke."));
    expect(await rejects(() => smt.applyPreset("input"))).toBe(
      "Something broke.",
    );

    helpers.failNextSync(hostError("GeneralException", "Something broke."));
    expect(await rejects(() => smt.applyNumberFormat("whole"))).toBe(
      "Something broke.",
    );
  });
});

// ---------------------------------------------------------------------------

// The capture's cell cap is counted over every area together, so two blocks
// that each fit can still be too much to hold - and the pane has to say so
// rather than imply a safety net that is not there.
describe("the undo budget over a ctrl-clicked selection", () => {
  it("paints both blocks and arms the note when they are over the cap", async () => {
    helpers.selectAreas(["Model!A1:A3000", "Model!C1:C3000"]);
    await smt.applyPreset("input");

    const theme = brand.deriveTheme(brand.DEFAULT_SETTINGS);
    expect(helpers.font("Model!A3000").color).toBe(theme.inputFont);
    expect(helpers.font("Model!C3000").color).toBe(theme.inputFont);
    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(true);
  });

  it("captures both blocks when they fit together", async () => {
    helpers.selectAreas(["Model!A1:A2000", "Model!C1:C2000"]);
    await smt.applyPreset("input");

    expect(smt.undoTarget()).toBe("Model!A1:A2000, Model!C1:C2000");
    expect(smt.lastUndoSkipped()).toBe(false);
    await smt.undoLastAction();
    expect(helpers.font("Model!C2000").color).toBe("#000000");
  });
});

// ---------------------------------------------------------------------------

// The pane's colours, font and currency live in module state, never in the
// workbook a format action reads: a brand set in the Brand tab has to reach
// every one of them without another round trip.
describe("a brand other than the shipped one", () => {
  const CUSTOM: SettingsModule.BrandSettings = {
    primary: "#7A0019",
    accent: "#FFC845",
    input: "#005EB8",
    formula: "#2B2B2B",
    link: "#0F7B3F",
    external: "#B00020",
    partial: "#6A1B9A",
    font: "Georgia",
    language: "en",
    currency: "$",
    autocolorOnEdit: false,
  };

  it("paints presets, fills and fonts from the active palette", async () => {
    brand.setActiveSettings(CUSTOM);
    const theme = brand.deriveTheme(CUSTOM);
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");

    await smt.applyPreset("title");
    expect(helpers.fill("Model!A1").color).toBe(theme.titleFill);
    expect(helpers.font("Model!A1").name).toBe("Georgia");
    // The preset left the brand primary on the cell, which is the fill cycle's
    // last rung: the next press clears rather than restarting the ladder.
    await smt.applyFillCycle();
    expect(helpers.fill("Model!A1").pattern).toBe("None");

    await smt.applyFillCycle();
    expect(helpers.fill("Model!A1").color).toBe(theme.headerFill);
    await smt.applyFontColorCycle();
    expect(helpers.font("Model!A1").color).toBe(theme.formulaFont);
  });

  it("cycles the currency format the brand asked for", async () => {
    brand.setActiveSettings(CUSTOM);
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");

    await smt.applyNumberFormat("currency");
    expect(helpers.numberFormat("Model!A1")).toContain("$");
    // The cycle reads that format back and steps on rather than restarting.
    await smt.applyNumberCycle("currency");
    expect(helpers.numberFormat("Model!A1")).toContain("#,##0.0");
  });

  it("steps the row-style cycle through the brand's own variants", async () => {
    brand.setActiveSettings(CUSTOM);
    const theme = brand.deriveTheme(CUSTOM);
    helpers.seed("Model!A1", [[1], [2]]);
    helpers.select("Model!A1:A2");

    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A2").color).toBe(theme.titleFill);
    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A2").color).toBe(theme.headerFill);
  });
});

// ---------------------------------------------------------------------------

// Independence: the formatting tools are the first thing a modeller touches,
// long before the Links tab has ever been opened, and a locked-down webview
// can refuse localStorage outright. Neither may stop a format action.
describe("formatting with no storage and no links registry", () => {
  const throwingStorage = {
    getItem: (): string => {
      throw new Error("storage blocked");
    },
    setItem: (): void => {
      throw new Error("storage blocked");
    },
    removeItem: (): void => {
      throw new Error("storage blocked");
    },
  };

  beforeEach(() => {
    vi.stubGlobal("localStorage", throwingStorage);
    helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!A1:B2");
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("runs every format action and touches no workbook setting", async () => {
    await smt.applyPreset("header");
    await smt.applyNumberFormat("percent");
    await smt.clearFormats();
    await smt.applyFillCycle();
    await smt.applyFontColorCycle();
    await smt.applyBorderCycle();
    await smt.applyRowStyleCycle("result");
    await smt.applyNumberCycle("percent");
    await smt.applyRowHeightCycle();
    await smt.applyColumnWidthCycle();
    await smt.undoLastAction();

    // The registry the Links tab keeps, and the auto-push switch beside it.
    expect(helpers.setting("PLSFIX_LINKS")).toBeNull();
    expect(helpers.setting("PLSFIX_AUTOPUSH")).toBeNull();
  });

  it("captures and paints a slot with storage refusing every call", async () => {
    helpers.setActiveCell("Model!A1");
    helpers.setFont("Model!A1", { name: "Georgia", bold: true });
    const slot = await smt.captureSlot(2);

    helpers.select("Model!D1:D2");
    await smt.applySlot(2, slot);
    expect(helpers.font("Model!D2").name).toBe("Georgia");
  });
});
