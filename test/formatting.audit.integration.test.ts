// Audit of the selection formatters against the strict fake host: what a
// protected sheet says, what a refused row-style cycle does to pls,fix Undo,
// the pre-1.9 single-area fallback, and the promise that formatting never
// depends on local storage or on the Links tab having been opened.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
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
