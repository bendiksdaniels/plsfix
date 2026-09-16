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
import type * as CyclesModule from "../src/cycles";

// Every host in this file answers like the real one: a scalar property read
// without a load() plus a context.sync() throws instead of quietly working.
enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;
let brand: typeof SettingsModule;
let cycles: typeof CyclesModule;
let theme: ReturnType<(typeof SettingsModule)["deriveTheme"]>;
let palette: (typeof SettingsModule)["DEFAULT_SETTINGS"];

// A fresh runtime plus a fresh module graph: the undo slot, the copy source, the
// overlay snapshots and the edit handler all start clean.
async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
  cycles = await import("../src/cycles");
  palette = brand.DEFAULT_SETTINGS;
  theme = brand.deriveTheme(palette);
}

// Reopening the file: same workbook model, brand new runtime and module state.
async function reopen(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ workbook, ...options });
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

// The gate that would have caught the borders gap: the whole cell map, byte for
// byte, before and after a mutator plus its undo.
async function expectExactUndo(
  run: () => Promise<unknown>,
  sheetName = "Model",
): Promise<void> {
  const before = helpers.cellMap(sheetName);
  await run();
  expect(helpers.cellMap(sheetName)).not.toEqual(before);
  await smt.undoLastAction();
  expect(helpers.cellMap(sheetName)).toEqual(before);
}

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------

describe("selection", () => {
  it("summarises a selection under the cap", async () => {
    helpers.seed("Model!A1", [
      [1, { formula: "=A1*2", value: 2 }, ""],
      ["#DIV/0!", { formula: "=Data!B1", value: 7 }, null],
    ]);
    helpers.select("Model!A1:C2");

    expect(await smt.inspectSelection()).toEqual({
      address: "Model!A1:C2",
      cells: 6,
      formulas: 2,
      errors: 1,
      blanks: 2,
    });
  });

  it("returns sentinels instead of reading a whole-column selection", async () => {
    helpers.select("Model!A:A");
    const summary = await smt.inspectSelection();

    expect(summary.address).toBe("Model!A:A");
    expect(summary.cells).toBe(1_048_576);
    expect(summary.formulas).toBe(-1);
    expect(summary.errors).toBe(-1);
    expect(summary.blanks).toBe(-1);
  });

  it("refuses a capped action on an oversized selection", async () => {
    helpers.select("Model!A1:A5001");
    expect(await rejects(() => smt.applyNumberFormat("whole"))).toBe(
      "Number formatting supports up to 5,000 selected cells at once.",
    );
  });

  it("parses sheet-qualified addresses, quotes and all", () => {
    expect(smt.parseAddress("Model!A1:C3")).toEqual({
      sheet: "Model",
      address: "A1:C3",
    });
    expect(smt.parseAddress("'Bob''s Model'!B2")).toEqual({
      sheet: "Bob's Model",
      address: "B2",
    });
    expect(smt.parseAddress("B2")).toEqual({ sheet: "", address: "B2" });
  });
});

// ---------------------------------------------------------------------------

describe("presets", () => {
  it("applies the title preset", async () => {
    helpers.select("Model!A1:C1");
    await smt.applyPreset("title");

    const font = helpers.font("Model!B1");
    expect(font).toMatchObject({
      name: palette.font,
      size: 15,
      bold: true,
      italic: false,
      color: theme.titleText,
    });
    expect(helpers.fill("Model!B1")).toEqual({
      color: theme.titleFill,
      pattern: "Solid",
      patternColor: "#FFFFFF",
    });
    expect(helpers.cell("Model!B1").horizontalAlignment).toBe("Left");
    expect(helpers.cell("Model!B1").verticalAlignment).toBe("Center");
    expect(helpers.rowHeight("Model", 0)).toBe(25);
  });

  it("puts the header rule on the bottom edge only", async () => {
    helpers.select("Model!A1:C3");
    await smt.applyPreset("header");

    expect(helpers.fill("Model!A1").color).toBe(theme.headerFill);
    expect(helpers.font("Model!A1").bold).toBe(true);
    expect(helpers.border("Model!A3", "bottom")).toEqual({
      style: "Continuous",
      color: theme.headerBorder,
      weight: "Thin",
    });
    expect(helpers.border("Model!A1", "bottom").style).toBe("None");
    expect(helpers.border("Model!A2", "bottom").style).toBe("None");
  });

  it("applies the input preset with no fill", async () => {
    helpers.setFill("Model!A1", { color: "#123456", pattern: "Solid" });
    helpers.select("Model!A1:B1");
    await smt.applyPreset("input");

    expect(helpers.font("Model!A1").color).toBe(theme.inputFont);
    expect(helpers.font("Model!A1").size).toBe(10);
    expect(helpers.fill("Model!A1")).toEqual({
      color: "#FFFFFF",
      pattern: "None",
      patternColor: "#FFFFFF",
    });
  });

  it("applies the formula preset", async () => {
    helpers.select("Model!A1:B1");
    await smt.applyPreset("formula");

    expect(helpers.font("Model!B1").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!B1").bold).toBe(false);
  });

  it("puts the result double rule on the top edge only", async () => {
    helpers.select("Model!A1:B3");
    await smt.applyPreset("result");

    expect(helpers.fill("Model!A2").color).toBe(theme.resultFill);
    expect(helpers.font("Model!A2").bold).toBe(true);
    expect(helpers.border("Model!A1", "top")).toEqual({
      style: "Double",
      color: theme.resultBorder,
      weight: "Thin",
    });
    expect(helpers.border("Model!A2", "top").style).toBe("None");
  });

  it("clears formats back to the host defaults", async () => {
    helpers.seed("Model!A1", [[1, 2]]);
    helpers.select("Model!A1:B1");
    await smt.applyPreset("title");
    await smt.clearFormats();

    expect(helpers.font("Model!A1")).toEqual({
      name: "Calibri",
      size: 11,
      bold: false,
      italic: false,
      color: "#000000",
      underline: "None",
    });
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.value("Model!A1")).toBe(1);
  });
});

// ---------------------------------------------------------------------------

describe("number formats", () => {
  it("applies every static format plus the branded currency", async () => {
    helpers.select("Model!A1:B2");

    await smt.applyNumberFormat("whole");
    expect(helpers.numberFormat("Model!B2")).toBe("#,##0;[Red](#,##0);-");
    await smt.applyNumberFormat("decimal");
    expect(helpers.numberFormat("Model!B2")).toBe("#,##0.0;[Red](#,##0.0);-");
    await smt.applyNumberFormat("percent");
    expect(helpers.numberFormat("Model!A1")).toBe("0.0%;[Red](0.0%);-");
    await smt.applyNumberFormat("currency");
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0 €;[Red](#,##0 €);-");
  });

  it("steps the general family three presses and wraps", async () => {
    helpers.select("Model!A1:B1");
    const family = cycles.buildNumberCycles(palette).general;

    for (const expected of [...family, family[0]]) {
      await smt.applyNumberCycle("general");
      expect(helpers.numberFormat("Model!A1")).toBe(expected);
    }
  });

  it("keeps stepping currency when the host rewrites the symbol", async () => {
    await boot({ rewriteCurrencyFormats: true });
    helpers.select("Model!A1:B1");
    const family = cycles.buildNumberCycles(palette).currency;

    for (const expected of family) {
      await smt.applyNumberCycle("currency");
      const stored = helpers.numberFormat("Model!A1");
      // The host handed back a locale-tagged code, not what we wrote.
      expect(stored).toContain("[$€-x-fake]");
      expect(cycles.canonicalNumberFormat(stored)).toBe(expected);
    }
  });

  it("steps decimals up and down", async () => {
    helpers.select("Model!A1:A1");
    await smt.applyNumberFormat("whole");

    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0.0;[Red](#,##0.0);-");
    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0.00;[Red](#,##0.00);-");
    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0.0;[Red](#,##0.0);-");
  });

  it("reads General as zero when adding the first decimal", async () => {
    helpers.select("Model!A1:A1");
    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!A1")).toBe("0.0");
  });
});

// ---------------------------------------------------------------------------

describe("style cycles", () => {
  it("draws a row-style border on every row of the selection", async () => {
    helpers.select("Model!A1:B3");
    await smt.applyRowStyleCycle("result");

    for (const address of ["Model!A1", "Model!A2", "Model!A3", "Model!B3"]) {
      expect(helpers.border(address, "top")).toEqual({
        style: "Double",
        color: theme.resultBorder,
        weight: "Thin",
      });
      expect(helpers.fill(address).color).toBe(theme.resultFill);
    }
  });

  it("steps the title cycle through every variant and wraps", async () => {
    helpers.select("Model!A1:C1");
    const variants = cycles.buildRowStyleCycles(palette).title;

    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A1").color).toBe(variants[0]?.fill);

    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A1").color).toBe(variants[1]?.fill);
    expect(helpers.border("Model!A1", "bottom").style).toBe("Continuous");

    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.font("Model!A1").color).toBe(palette.accent);

    await smt.applyRowStyleCycle("title");
    expect(helpers.fill("Model!A1").color).toBe(variants[0]?.fill);
  });

  // BUG: excel.ts:465 caps the per-row loop at 100 rows. Above that the code
  // falls back to range-level edge borders, which is the exact defect the
  // per-row loop exists to avoid (tasks/lessons.md, 2026-08-27) — a 101-row
  // selection gets its double rule on row 1 only, silently.
  it("draws a row-style border on every row above a hundred rows", async () => {
    helpers.select("Model!A1:A101");
    await smt.applyRowStyleCycle("result");

    expect(helpers.border("Model!A1", "top").style).toBe("Double");
    expect(helpers.border("Model!A50", "top").style).toBe("Double");
    expect(helpers.border("Model!A101", "top").style).toBe("Double");
  });

  it("steps the item cycle", async () => {
    helpers.select("Model!A1:A2");
    await smt.applyRowStyleCycle("item");
    expect(helpers.fill("Model!A1").pattern).toBe("None");

    await smt.applyRowStyleCycle("item");
    expect(helpers.fill("Model!A1").color).toBe(theme.headerFill);
  });

  it("cycles fills and lands back on no fill", async () => {
    helpers.select("Model!A1:B2");
    const cycle = cycles.buildFillCycle(palette);

    for (const step of cycle.slice(0, 4)) {
      await smt.applyFillCycle();
      expect(helpers.fill("Model!B2").color).toBe(step);
      expect(helpers.fill("Model!B2").pattern).toBe("Solid");
    }

    await smt.applyFillCycle();
    expect(helpers.fill("Model!B2")).toEqual({
      color: "#FFFFFF",
      pattern: "None",
      patternColor: "#FFFFFF",
    });

    await smt.applyFillCycle();
    expect(helpers.fill("Model!B2").color).toBe(cycle[0]);
  });

  it("cycles font colours and wraps", async () => {
    helpers.select("Model!A1:A1");
    const cycle = cycles.buildFontCycle(palette);

    for (const step of cycle) {
      await smt.applyFontColorCycle();
      expect(helpers.font("Model!A1").color).toBe(step);
    }
    await smt.applyFontColorCycle();
    expect(helpers.font("Model!A1").color).toBe(cycle[0]);
  });

  // The border cycle draws on the selection's own edges, so the read-back is a
  // map of cells: the outer band carries the box, the inner cells the grid.
  describe("borders", () => {
    const drawn = (weight = "Thin") => ({
      style: "Continuous",
      color: palette.primary,
      weight,
    });
    const none = (address: string, edge: "top" | "bottom" | "left" | "right") =>
      helpers.border(address, edge).style;

    beforeEach(() => {
      helpers.select("Model!B2:D4");
    });

    it("rules the bottom of the selection, then makes it a total", async () => {
      await smt.applyBorderCycle();
      for (const address of ["Model!B4", "Model!C4", "Model!D4"]) {
        expect(helpers.border(address, "bottom")).toEqual(drawn());
      }
      expect(none("Model!B2", "top")).toBe("None");
      expect(none("Model!B2", "bottom")).toBe("None");

      await smt.applyBorderCycle();
      expect(helpers.border("Model!C4", "bottom")).toEqual(drawn("Medium"));
    });

    it("draws the result line, then a box, then the grid, then clears", async () => {
      await smt.applyBorderCycle();
      await smt.applyBorderCycle();

      await smt.applyBorderCycle();
      expect(helpers.border("Model!C2", "top")).toEqual(drawn());
      // A double rule keeps Excel's own weight, so only style and colour.
      const result = helpers.border("Model!C4", "bottom");
      expect(result.style).toBe("Double");
      expect(result.color).toBe(palette.primary);

      await smt.applyBorderCycle();
      expect(helpers.border("Model!D2", "top")).toEqual(drawn());
      expect(helpers.border("Model!D4", "bottom")).toEqual(drawn());
      expect(helpers.border("Model!B3", "left")).toEqual(drawn());
      expect(helpers.border("Model!D3", "right")).toEqual(drawn());
      // A box has nothing inside it.
      expect(none("Model!C3", "bottom")).toBe("None");
      expect(none("Model!C3", "right")).toBe("None");

      await smt.applyBorderCycle();
      expect(helpers.border("Model!C3", "bottom")).toEqual(drawn());
      expect(helpers.border("Model!C3", "right")).toEqual(drawn());
      expect(helpers.border("Model!B2", "top")).toEqual(drawn());
      expect(helpers.border("Model!D4", "right")).toEqual(drawn());

      await smt.applyBorderCycle();
      for (const address of ["Model!B2", "Model!C3", "Model!D4"]) {
        for (const edge of ["top", "bottom", "left", "right"] as const) {
          expect(none(address, edge)).toBe("None");
        }
      }
    });

    it("puts the borders it overwrote back", async () => {
      await expectExactUndo(() => smt.applyBorderCycle());
    });

    it("boxes a single cell, which has no lines inside to draw", async () => {
      helpers.select("Model!A1");
      for (let press = 0; press < 4; press += 1) await smt.applyBorderCycle();

      for (const edge of ["top", "bottom", "left", "right"] as const) {
        expect(helpers.border("Model!A1", edge)).toEqual(drawn());
      }

      // A grid would look exactly like the box, so the cycle comes home instead.
      await smt.applyBorderCycle();
      for (const edge of ["top", "bottom", "left", "right"] as const) {
        expect(none("Model!A1", edge)).toBe("None");
      }
    });
  });

  // Sizes are sheet state, not cell state: the first row or column of the
  // selection carries the rung, the whole band is written, and pls,fix Undo can
  // reach neither.
  describe("sizes", () => {
    it("steps the selected rows through the height ladder and wraps", async () => {
      helpers.select("Model!B2:C4");
      const rungs = cycles.buildSizeCycles().rowHeight.length;
      const walked: number[] = [];

      for (let press = 0; press < rungs; press += 1) {
        await smt.applyRowHeightCycle();
        // The whole band moves together, first row to last.
        expect(helpers.rowHeight("Model", 3)).toBe(
          helpers.rowHeight("Model", 1),
        );
        walked.push(helpers.rowHeight("Model", 1));
      }

      expect(walked).toEqual([18, 21, 24, 30, 15]);
    });

    it("widens the selected columns and comes back narrow", async () => {
      helpers.select("Model!B2:C4");
      const rungs = cycles.buildSizeCycles().columnWidth.length;
      const walked: number[] = [];

      for (let press = 0; press < rungs; press += 1) {
        await smt.applyColumnWidthCycle();
        walked.push(helpers.columnWidth("Model", 2));
      }

      expect(walked).toEqual([80, 96, 120, 48, 64]);
    });

    it("leaves the rows and columns outside the selection alone", async () => {
      helpers.select("Model!B2:C4");
      await smt.applyRowHeightCycle();
      await smt.applyColumnWidthCycle();

      expect(helpers.rowHeight("Model", 1)).toBe(18);
      expect(helpers.rowHeight("Model", 4)).toBe(15);
      expect(helpers.columnWidth("Model", 1)).toBe(80);
      expect(helpers.columnWidth("Model", 3)).toBe(64);
    });

    it("reads the rung off the first row of the selection", async () => {
      helpers.sheet("Model").rowHeights.set(1, 21);
      helpers.select("Model!B2:C4");
      await smt.applyRowHeightCycle();

      for (const row of [1, 2, 3]) {
        expect(helpers.rowHeight("Model", row)).toBe(24);
      }
    });

    // Clicking the column headers is how a modeller reaches this: a cell cap
    // would refuse the selection the action is made for.
    it("takes a whole-column selection", async () => {
      helpers.select("Model!B:C");
      await smt.applyColumnWidthCycle();

      expect(helpers.columnWidth("Model", 1)).toBe(80);
      expect(helpers.columnWidth("Model", 2)).toBe(80);
      expect(helpers.columnWidth("Model", 3)).toBe(64);
    });

    it("runs outside pls,fix Undo, which cannot reach sheet state", async () => {
      helpers.select("Model!A1:C1");
      await smt.applyPreset("title");
      expect(helpers.rowHeight("Model", 0)).toBe(25);

      // The preset's 25 pt is no rung of ours, so the ladder starts at its foot.
      await smt.applyRowHeightCycle();
      expect(helpers.rowHeight("Model", 0)).toBe(15);
      // The cycle captured nothing, so the slot still holds the preset.
      expect(smt.undoTarget()).toBe("Model!A1:C1");

      await smt.undoLastAction();
      expect(helpers.rowHeight("Model", 0)).toBe(15);
    });
  });
});

// ---------------------------------------------------------------------------

describe("autocolor", () => {
  const grid = [
    [1234, { formula: "=A1+B1", value: 3 }],
    [
      { formula: "=Data!B4", value: 9 },
      { formula: "=[Model.xlsx]Sheet1!A1", value: 4 },
    ],
    [{ formula: "=A1*1.05", value: 5 }, ""],
  ];

  it("colours every class from the active theme", async () => {
    helpers.seed("Model!A1", grid);
    helpers.select("Model!A1:B3");
    await smt.autocolorSelection();

    expect(helpers.font("Model!A1").color).toBe(theme.inputFont);
    expect(helpers.font("Model!B1").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!A2").color).toBe(theme.linkFont);
    expect(helpers.font("Model!B2").color).toBe(theme.externalFont);
    expect(helpers.font("Model!A3").color).toBe(theme.partialFont);
    // Blank cells keep whatever the modeller had.
    expect(helpers.font("Model!B3").color).toBe("#000000");
  });

  it("leaves a text label's own colour alone but still colours a number", async () => {
    helpers.seed("Model!A1", [["Revenue", 1234]]);
    // A navy title font, exactly what a header preset would have left behind.
    helpers.setFont("Model!A1", { color: theme.titleText });
    helpers.select("Model!A1:B1");
    await smt.autocolorSelection();

    expect(helpers.font("Model!A1").color).toBe(theme.titleText);
    expect(helpers.font("Model!B1").color).toBe(theme.inputFont);
  });

  it("refuses an oversized autocolor", async () => {
    helpers.select("Model!A1:A5001");
    expect(await rejects(() => smt.autocolorSelection())).toBe(
      "Autocolor supports up to 5,000 selected cells at once.",
    );
  });

  it("writes the colour key block", async () => {
    helpers.select("Model!B2");
    await smt.insertColorKey();

    expect(helpers.value("Model!B2")).toBe("Color key");
    expect(helpers.fill("Model!B2").color).toBe(theme.titleFill);
    expect(helpers.font("Model!B2")).toMatchObject({
      color: theme.titleText,
      bold: true,
      size: 11,
    });

    expect(helpers.value("Model!B3")).toBe("Hardcoded input");
    expect(helpers.value("Model!C3")).toBe(1234);
    expect(helpers.numberFormat("Model!C3")).toBe("#,##0");
    expect(helpers.font("Model!B3").color).toBe(theme.inputFont);

    expect(helpers.value("Model!C4")).toBe("=A1+B1");
    // Text format first, so the example formulas never compute.
    expect(helpers.numberFormat("Model!C4")).toBe("@");
    expect(helpers.font("Model!B4").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!B5").color).toBe(theme.linkFont);
    expect(helpers.font("Model!B6").color).toBe(theme.externalFont);
    expect(helpers.font("Model!B7").color).toBe(theme.partialFont);
    expect(helpers.fill("Model!B7").pattern).toBe("None");
  });
});

// ---------------------------------------------------------------------------

describe("audit overlay", () => {
  // R1C1 makes a filled formula read identically everywhere, so equality with a
  // neighbour is the consistency test. This block yields every mark.
  const r1c1 = [
    ["=R[1]C", "=R[1]C", "=R[1]C"],
    ["=R[1]C", "=RC[-1]", 5],
    ["=X", "", ""],
  ];

  function seedBlock(): void {
    helpers.seed(
      "Model!A1",
      r1c1.map((row) =>
        row.map((entry) =>
          typeof entry === "string" && entry.startsWith("=")
            ? { formula: entry, r1c1: entry, value: 1 }
            : entry,
        ),
      ),
    );
  }

  function striped(patternColor: string) {
    return { color: "#FFFFFF", pattern: "CrissCross", patternColor };
  }

  it("paints one pattern per audit mark", async () => {
    seedBlock();
    helpers.select("Model!A1:C3");
    expect(await smt.toggleAuditOverlay()).toBe(true);

    const patternColor = brand.tint(palette.primary, 0.55);
    expect(helpers.fill("Model!A1")).toEqual(striped(patternColor));
    expect(helpers.fill("Model!B1").pattern).toBe("LightHorizontal");
    expect(helpers.fill("Model!C1").pattern).toBe("LightHorizontal");
    expect(helpers.fill("Model!A2").pattern).toBe("LightVertical");
    expect(helpers.fill("Model!B2")).toEqual({
      color: "#E8B4B4",
      pattern: "Solid",
      patternColor: "#E8B4B4",
    });
    expect(helpers.fill("Model!A3").pattern).toBe("Solid");
    // A plain value beside a formula reads as a typed hardcode, not a
    // deviation: the same solid tint the stripes carry as their pattern
    // colour, not the lone fill's red.
    expect(helpers.fill("Model!C2")).toEqual({
      color: patternColor,
      pattern: "Solid",
      patternColor,
    });
  });

  it("restores the modeller's own striped fill byte for byte", async () => {
    seedBlock();
    helpers.setFill("Model!A1", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    const before = helpers.cellMap("Model");

    helpers.select("Model!A1:C3");
    await smt.toggleAuditOverlay();
    expect(helpers.fill("Model!A1").pattern).toBe("CrissCross");

    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(helpers.fill("Model!A1")).toEqual({
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("stores the snapshot in workbook settings and clears it on restore", async () => {
    seedBlock();
    helpers.select("Model!A1:C3");

    await smt.toggleAuditOverlay();
    const stored = workbook.settings.get("smtAuditOverlay") ?? "";
    expect(JSON.parse(stored)).toHaveLength(1);
    expect(JSON.parse(stored)[0]).toMatchObject({ address: "A1:C3" });

    await smt.toggleAuditOverlay();
    expect(workbook.settings.get("smtAuditOverlay")).toBe("");
  });

  it("treats a single cell as the block around it", async () => {
    seedBlock();
    helpers.select("Model!B2");
    await smt.toggleAuditOverlay();

    // The whole A1:C3 region was painted, not just B2.
    expect(helpers.fill("Model!A1").pattern).toBe("CrissCross");
    const stored = JSON.parse(workbook.settings.get("smtAuditOverlay") ?? "[]");
    expect(stored[0].address).toBe("A1:C3");
  });

  it("refuses an oversized overlay", async () => {
    helpers.select("Model!A1:A5001");
    expect(await rejects(() => smt.toggleAuditOverlay())).toBe(
      "The audit overlay supports up to 5,000 cells at once.",
    );
  });

  it("restores last session's fills after the workbook is reopened", async () => {
    seedBlock();
    helpers.setFill("Model!A1", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    const before = helpers.cellMap("Model");
    helpers.select("Model!A1:C3");
    await smt.toggleAuditOverlay();

    // The runtime dies with the pane; the snapshot rides along in the file.
    await reopen();
    expect(await smt.restorePersistedOverlay()).toBe(true);
    expect(helpers.cellMap("Model")).toEqual(before);
    expect(workbook.settings.get("smtAuditOverlay")).toBe("");
  });

  it("reports nothing to restore when no overlay was saved", async () => {
    expect(await smt.restorePersistedOverlay()).toBe(false);
  });

  it("snapshots and restores fills on demand", async () => {
    helpers.setFill("Model!A1:B1", {
      color: "#112233",
      pattern: "Grid",
      patternColor: "#445566",
    });
    helpers.select("Model!A1:B1");

    const host = globalThis as unknown as {
      Excel: { run: (cb: (context: never) => unknown) => Promise<unknown> };
    };
    await host.Excel.run(async (context: never) => {
      const range = (
        context as unknown as {
          workbook: { getSelectedRange: () => unknown };
        }
      ).workbook.getSelectedRange();
      await smt.snapshotFills(context, range as Excel.Range);
    });

    helpers.setFill("Model!A1:B1", { color: "#FF0000", pattern: "Solid" });
    await host.Excel.run(async (context: never) => {
      await smt.restoreFills(context);
    });

    expect(helpers.fill("Model!B1")).toEqual({
      color: "#112233",
      pattern: "Grid",
      patternColor: "#445566",
    });
  });
});

// ---------------------------------------------------------------------------

describe("pls,fix undo", () => {
  function seedModel(): void {
    helpers.seed("Model!A1", [
      [1, 2, 3],
      [{ formula: "=A1*2", value: 2 }, 5, 6],
      [7, 8, { formula: "=IFERROR(A1,0)", value: 1 }],
    ]);
    helpers.setNumberFormat("Model!A1:C3", "#,##0");
    helpers.setFill("Model!B2", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    helpers.setFont("Model!C1", { bold: true, italic: true, color: "#123456" });
    helpers.select("Model!A1:C3");
  }

  beforeEach(seedModel);

  it("restores a preset exactly", async () => {
    await expectExactUndo(() => smt.applyPreset("title"));
  });

  it("restores cleared formats exactly", async () => {
    await expectExactUndo(() => smt.clearFormats());
  });

  it("restores a number format exactly", async () => {
    await expectExactUndo(() => smt.applyNumberFormat("currency"));
  });

  it("restores a number cycle exactly", async () => {
    await expectExactUndo(() => smt.applyNumberCycle("percent"));
  });

  it("restores a row style cycle exactly, borders and all", async () => {
    await expectExactUndo(() => smt.applyRowStyleCycle("result"));
  });

  it("restores a fill cycle exactly", async () => {
    await expectExactUndo(() => smt.applyFillCycle());
  });

  it("restores a font colour cycle exactly", async () => {
    await expectExactUndo(() => smt.applyFontColorCycle());
  });

  it("restores autocolor exactly", async () => {
    await expectExactUndo(() => smt.autocolorSelection());
  });

  it("restores a scale exactly", async () => {
    await expectExactUndo(() => smt.scaleSelection(1000));
  });

  it("restores a sign flip exactly", async () => {
    await expectExactUndo(() => smt.applySignFlip());
  });

  it("restores a decimal step exactly", async () => {
    await expectExactUndo(() => smt.applyDecimalStep(1));
  });

  it("restores an IFERROR toggle exactly", async () => {
    await expectExactUndo(() => smt.toggleIfErrorGuard());
  });

  it("restores a paste exactly, including the grown footprint", async () => {
    helpers.select("Model!A1:C2");
    await smt.markCopySource();
    helpers.select("Model!A3");
    await expectExactUndo(() => smt.pasteSpecial("values"));
  });

  it("restores a colour key exactly", async () => {
    helpers.select("Data!B2");
    await expectExactUndo(() => smt.insertColorKey(), "Data");
  });

  it("restores an inserted CAGR exactly", async () => {
    helpers.select("Model!A1:C1");
    await expectExactUndo(() => smt.insertCagr());
  });

  it("restores a fast fill exactly", async () => {
    helpers.select("Model!A2");
    await expectExactUndo(() => smt.fastFillAuto("right"));
  });

  it("names the range it will restore", async () => {
    expect(smt.undoTarget()).toBeNull();
    await smt.applyPreset("input");
    expect(smt.undoTarget()).toBe("Model!A1:C3");
    expect(smt.lastUndoSkipped()).toBe(false);

    await smt.undoLastAction();
    expect(smt.undoTarget()).toBeNull();
  });

  it("runs an oversized action but says the safety net is off", async () => {
    helpers.select("Model!A1:A5001");
    await smt.applyPreset("input");

    expect(helpers.font("Model!A5001").color).toBe(theme.inputFont);
    expect(smt.undoTarget()).toBeNull();
    // Read-once: a later action must not inherit the flag.
    expect(smt.lastUndoSkipped()).toBe(true);
    expect(smt.lastUndoSkipped()).toBe(false);
  });

  it("refuses to undo when there is nothing to undo", async () => {
    expect(await rejects(() => smt.undoLastAction())).toBe(
      "There is no pls,fix action to undo yet.",
    );
  });

  it("reports a deleted sheet instead of restoring into thin air", async () => {
    helpers.select("Data!A1:B2");
    await smt.applyPreset("header");
    helpers.deleteSheet("Data");

    expect(await rejects(() => smt.undoLastAction())).toBe(
      "The sheet that action ran on is gone.",
    );
    expect(smt.undoTarget()).toBeNull();
  });

  it("cannot reach row height, which is sheet state", async () => {
    helpers.select("Model!A1:C1");
    await smt.applyPreset("title");
    expect(helpers.rowHeight("Model", 0)).toBe(25);

    await smt.undoLastAction();
    // Documented limitation: getCellProperties carries no row height.
    expect(helpers.rowHeight("Model", 0)).toBe(25);
  });

  it("stacks five actions and undoes them newest first, content and all", async () => {
    const addresses = [
      "Model!A5",
      "Model!B5",
      "Model!C5",
      "Model!D5",
      "Model!E5",
    ];
    const snapshots = [helpers.cellMap("Model")];
    for (const address of addresses) {
      helpers.select(address);
      await smt.applyPreset("input");
      snapshots.push(helpers.cellMap("Model"));
    }

    for (let i = addresses.length; i >= 1; i--) {
      expect(smt.undoTarget()).toBe(addresses[i - 1]);
      await smt.undoLastAction();
      expect(helpers.cellMap("Model")).toEqual(snapshots[i - 1]);
    }
    expect(smt.undoTarget()).toBeNull();
  });

  it("drops the oldest entry once a sixth action is captured", async () => {
    const addresses = [
      "Model!A5",
      "Model!B5",
      "Model!C5",
      "Model!D5",
      "Model!E5",
      "Model!F5",
    ];
    for (const address of addresses) {
      helpers.select(address);
      await smt.applyPreset("input");
    }

    // Only the newest five survive; A5's capture is gone for good.
    for (const address of [...addresses].slice(1).reverse()) {
      expect(smt.undoTarget()).toBe(address);
      await smt.undoLastAction();
    }
    expect(smt.undoTarget()).toBeNull();
    expect(await rejects(() => smt.undoLastAction())).toBe(
      "There is no pls,fix action to undo yet.",
    );
  });

  it("keeps the stack within the 25,000-cell budget", async () => {
    // Six captures at the 5,000-cell per-action cap: the sixth pushes the
    // running total past UNDO_CELL_BUDGET (25,000 = 5 x the per-action cap),
    // so the oldest has to go the same way it would on depth alone. The
    // budget rule is exercised on its own, independent of depth, in
    // undo-stack.test.ts.
    const addresses = [
      "Model!A10:B2509",
      "Model!D10:E2509",
      "Model!G10:H2509",
      "Model!J10:K2509",
      "Model!M10:N2509",
      "Model!P10:Q2509",
    ];
    for (const address of addresses) {
      helpers.select(address);
      await smt.applyPreset("input");
    }

    for (const address of [...addresses].slice(1).reverse()) {
      expect(smt.undoTarget()).toBe(address);
      await smt.undoLastAction();
    }
    expect(smt.undoTarget()).toBeNull();
  });

  it("keeps the snapshot when the restore fails, so a retry can succeed", async () => {
    await smt.applyPreset("input"); // captures Model!A1:C3, seedModel's selection
    const changed = helpers.cellMap("Model");

    helpers.failNextSync();
    await expect(smt.undoLastAction()).rejects.toThrow();
    expect(smt.undoTarget()).toBe("Model!A1:C3");
    expect(helpers.cellMap("Model")).toEqual(changed);

    // Retried without a forced failure, the same entry now restores.
    await smt.undoLastAction();
    expect(smt.undoTarget()).toBeNull();
  });

  it("says how many actions are left to undo, then says there are none", async () => {
    helpers.select("Model!A5");
    await smt.applyPreset("input");
    helpers.select("Model!B5");
    await smt.applyPreset("input");
    helpers.select("Model!C5");
    await smt.applyPreset("input");

    expect(await smt.undoLastAction()).toBe(
      "Undone: Model!C5. 2 more to undo.",
    );
    expect(await smt.undoLastAction()).toBe(
      "Undone: Model!B5. 1 more to undo.",
    );
    expect(await smt.undoLastAction()).toBe(
      "Undone: Model!A5. Nothing more to undo.",
    );
  });

  // Excel for Mac (16.107) answers an unfilled cell's fill with a null
  // pattern and an empty patternColor; feeding that straight back into
  // setCellProperties was refused with a raw InvalidArgument error, leaving
  // the entry stuck on the stack even though the values had already
  // restored. seedModel's selection is exactly this mix: most of A1:C3 was
  // never filled, and B2 carries a real striped pattern with its own colours.
  it("restores a block over unfilled cells after a fill-painting action, no error", async () => {
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    const originalB2 = helpers.fill("Model!B2");
    expect(originalB2).toEqual({
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });

    // The title preset clears every cell's fill and paints it solid navy,
    // overwriting both the unfilled cells and B2's own pattern.
    await smt.applyPreset("title");
    expect(helpers.fill("Model!A1").pattern).not.toBe("None");
    expect(helpers.fill("Model!B2").pattern).not.toBe("LightUp");

    // The restore does not throw, and the entry is consumed (not left
    // retryable the way a genuinely failed restore stays).
    await expect(smt.undoLastAction()).resolves.toBe(
      "Undone: Model!A1:C3. Nothing more to undo.",
    );
    expect(smt.undoTarget()).toBeNull();

    // The originally unfilled cell is unfilled again. Its pattern is the
    // observable fact; a fill that never renders carries no colour promise
    // (Excel itself leaves a stale colour under a None pattern untouched
    // when only the pattern is written back, so settableFill sends "None"
    // alone rather than a colour nobody can see).
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    // The originally filled cell's colour is back...
    expect(helpers.fill("Model!B2").color).toBe(originalB2.color);
    // ...and so is its pattern colour, not just the fill colour.
    expect(helpers.fill("Model!B2")).toEqual(originalB2);
  });
});

// ---------------------------------------------------------------------------

describe("copy and paste", () => {
  beforeEach(() => {
    helpers.seed("Model!A1", [
      [1, { formula: "=A1+1", value: 2 }],
      [3, { formula: "=A2+1", value: 4 }],
    ]);
    helpers.setFill("Model!A1:B2", { color: "#EEDDCC", pattern: "Solid" });
    helpers.setNumberFormat("Model!A1:B2", "#,##0.0");
  });

  it("marks and reports the copy source", async () => {
    expect(smt.copySourceLabel()).toBeNull();
    helpers.select("Model!A1:B2");
    expect(await smt.markCopySource()).toBe("Model!A1:B2");
    expect(smt.copySourceLabel()).toBe("Model!A1:B2");
  });

  it("pastes values without carrying the formats", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteSpecial("values");

    expect(helpers.value("Data!B1")).toBe(2);
    expect(helpers.formula("Data!B1")).toBe(2);
    expect(helpers.fill("Data!B1").pattern).toBe("None");
    expect(helpers.numberFormat("Data!B1")).toBe("General");
  });

  it("pastes formats without carrying the values", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteSpecial("formats");

    expect(helpers.fill("Data!B2").color).toBe("#EEDDCC");
    expect(helpers.numberFormat("Data!B2")).toBe("#,##0.0");
    expect(helpers.value("Data!B2")).toBe("");
  });

  it("transposes a row into a column", async () => {
    helpers.seed("Model!D1", [[10, 20, 30]]);
    helpers.select("Model!D1:F1");
    await smt.markCopySource();
    helpers.select("Data!A1");
    await smt.pasteSpecial("transpose");

    expect(helpers.value("Data!A1")).toBe(10);
    expect(helpers.value("Data!A2")).toBe(20);
    expect(helpers.value("Data!A3")).toBe(30);
  });

  it("keeps formula text byte for byte", async () => {
    helpers.select("Model!A1:B2");
    await smt.markCopySource();
    helpers.select("Data!C5");
    await smt.pastePreserveFormulas();

    expect(helpers.formula("Data!D5")).toBe("=A1+1");
    expect(helpers.formula("Data!D6")).toBe("=A2+1");
    expect(helpers.formula("Data!C5")).toBe(1);
  });

  it("asks for a source before pasting", async () => {
    helpers.select("Data!A1");
    expect(await rejects(() => smt.pasteSpecial("values"))).toBe(
      "Mark a copy source first.",
    );
  });

  it("reports a deleted source sheet and forgets it", async () => {
    helpers.select("Data!A1:B1");
    await smt.markCopySource();
    helpers.deleteSheet("Data");
    helpers.select("Model!D1");

    expect(await rejects(() => smt.pasteSpecial("values"))).toBe(
      "The copy source sheet is gone. Mark a new source.",
    );
    expect(smt.copySourceLabel()).toBeNull();
  });
});

// ---------------------------------------------------------------------------

describe("fast fill", () => {
  it("fills right as far as the header row runs", async () => {
    helpers.seed("Model!B1", [["Q1", "Q2", "Q3", "Q4"]]);
    helpers.seed("Model!B2", [[{ formula: "=A2*2", value: 8 }]]);
    helpers.select("Model!B2");

    await smt.fastFillAuto("right");

    expect(helpers.formula("Model!C2")).toBe("=A2*2");
    expect(helpers.formula("Model!E2")).toBe("=A2*2");
    expect(helpers.formula("Model!F2")).toBe("");
  });

  it("fills down as far as the neighbour column runs", async () => {
    helpers.seed("Model!A2", [
      ["Rent"],
      ["Wages"],
      ["Fuel"],
      ["Other"],
      ["Tax"],
    ]);
    helpers.seed("Model!B2", [[{ formula: "=B1*2", value: 4 }]]);
    helpers.select("Model!B2");

    await smt.fastFillAuto("down");

    expect(helpers.formula("Model!B6")).toBe("=B1*2");
    expect(helpers.formula("Model!B7")).toBe("");
  });

  it("fills down from column A without walking off the sheet", async () => {
    helpers.seed("Model!B2", [["one"], ["two"], ["three"]]);
    helpers.seed("Model!A2", [[{ formula: "=A1*2", value: 2 }]]);
    helpers.select("Model!A2");

    await smt.fastFillAuto("down");

    expect(helpers.formula("Model!A4")).toBe("=A1*2");
    expect(helpers.formula("Model!A5")).toBe("");
  });

  it("carries the source cell's number format when it fills right", async () => {
    helpers.seed("Model!B1", [["Q1", "Q2", "Q3", "Q4"]]);
    helpers.seed("Model!B2", [[{ formula: "=A2*2", value: 8 }]]);
    helpers.setNumberFormat("Model!B2", "#,##0.0");
    helpers.select("Model!B2");

    await smt.fastFillAuto("right");

    expect(helpers.numberFormat("Model!C2")).toBe("#,##0.0");
    expect(helpers.numberFormat("Model!E2")).toBe("#,##0.0");
    // The header row the fill sized itself against is untouched.
    expect(helpers.numberFormat("Model!C1")).toBe("General");
  });

  it("carries the source cell's number format when it fills down", async () => {
    helpers.seed("Model!A2", [
      ["Rent"],
      ["Wages"],
      ["Fuel"],
      ["Other"],
      ["Tax"],
    ]);
    helpers.seed("Model!B2", [[{ formula: "=B1*2", value: 4 }]]);
    helpers.setNumberFormat("Model!B2", "0.00%");
    helpers.select("Model!B2");

    await smt.fastFillAuto("down");

    expect(helpers.numberFormat("Model!B6")).toBe("0.00%");
    // The label column beside it is untouched.
    expect(helpers.numberFormat("Model!A4")).toBe("General");
  });

  it("refuses when there is no neighbour data to size the fill", async () => {
    helpers.seed("Model!C5", [[{ formula: "=C4*2", value: 1 }]]);
    helpers.select("Model!C5");

    expect(await rejects(() => smt.fastFillAuto("down"))).toBe(
      "No neighbor data to size the fill.",
    );
  });

  it("refuses when the active cell holds no formula", async () => {
    helpers.seed("Model!B2", [[42]]);
    helpers.select("Model!B2");

    expect(await rejects(() => smt.fastFillAuto("right"))).toBe(
      "The active cell must contain a formula.",
    );
  });
});

// ---------------------------------------------------------------------------

describe("formula edits", () => {
  it("writes a CAGR just past a row of periods", async () => {
    helpers.seed("Model!A2", [[100, 110, 121, 133.1]]);
    helpers.select("Model!A2:D2");
    await smt.insertCagr();

    expect(helpers.formula("Model!E2")).toBe("=(D2/A2)^(1/3)-1");
    expect(helpers.numberFormat("Model!E2")).toBe("0.0%;[Red](0.0%);-");
  });

  it("writes a CAGR just below a column of periods", async () => {
    helpers.seed("Model!B1", [[100], [110], [121]]);
    helpers.select("Model!B1:B3");
    await smt.insertCagr();

    expect(helpers.formula("Model!B4")).toBe("=(B3/B1)^(1/2)-1");
  });

  it("refuses a CAGR on the wrong shape", async () => {
    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.insertCagr())).toBe(
      "Select one row or column with at least two periods.",
    );

    helpers.select("Model!A1");
    expect(await rejects(() => smt.insertCagr())).toBe(
      "Select one row or column with at least two periods.",
    );
  });

  it("toggles the IFERROR guard on and back off", async () => {
    helpers.seed("Model!A1", [[{ formula: "=A9/B9", value: 1 }, 5]]);
    helpers.select("Model!A1:B1");

    await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A1")).toBe("=IFERROR(A9/B9,0)");
    expect(helpers.formula("Model!B1")).toBe(5);

    await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A1")).toBe("=A9/B9");
  });

  it("scales numbers and wraps formulas", async () => {
    helpers.seed("Model!A1", [[5, { formula: "=B9", value: 2 }]]);
    helpers.select("Model!A1:B1");

    await smt.scaleSelection(1000);
    expect(helpers.value("Model!A1")).toBe(5000);
    expect(helpers.formula("Model!B1")).toBe("=(B9)*1000");

    await smt.scaleSelection(0.001);
    expect(helpers.value("Model!A1")).toBe(5);
    expect(helpers.formula("Model!B1")).toBe("=((B9)*1000)/1000");
  });

  it("flips signs and unwraps its own wrap", async () => {
    helpers.seed("Model!A1", [[7, { formula: "=B9+C9", value: 3 }]]);
    helpers.select("Model!A1:B1");

    await smt.applySignFlip();
    expect(helpers.value("Model!A1")).toBe(-7);
    expect(helpers.formula("Model!B1")).toBe("=-(B9+C9)");

    await smt.applySignFlip();
    expect(helpers.value("Model!A1")).toBe(7);
    expect(helpers.formula("Model!B1")).toBe("=B9+C9");
  });
});

// ---------------------------------------------------------------------------

describe("charts", () => {
  function seedBridge(close: number): void {
    helpers.seed("Model!A1", [["EBITDA bridge"]]);
    helpers.seed("Model!A2", [
      ["Opening", 100],
      ["Price", 20],
      ["Cost", -30],
      ["Closing", close],
    ]);
    helpers.select("Model!A2:B5");
  }

  it("keeps the waterfall when the host refuses the chart surface", async () => {
    // Excel for the web: font and corners are UnsupportedOperation on chartex
    // charts. The batch that carries them fails on its own; the chart, its
    // title, axes and branded points still land.
    await boot({ chartSurfaceUnsupported: true });
    seedBridge(90);
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90",
    );
    const chart = workbook.charts[0];
    expect(chart).toMatchObject({
      chartType: "Waterfall",
      title: "EBITDA bridge",
    });
    expect(chart?.font).toEqual({});
    expect(chart?.roundedCorners).toBeUndefined();
    expect(chart?.axes.value.fontName).toBe(palette.font);
  });

  it("adds a native waterfall and reports a clean reconciliation", async () => {
    seedBridge(90);
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90",
    );

    const chart = workbook.charts[0];
    expect(chart).toMatchObject({
      chartType: "Waterfall",
      sourceAddress: "Model!A2:B5",
      seriesBy: "Auto",
      title: "EBITDA bridge",
    });
    expect(chart?.font.name).toBe(palette.font);
    expect(chart?.titleFont).toMatchObject({
      bold: true,
      color: palette.primary,
    });
    expect(chart?.legend.visible).toBe(false);
    // Just the values: every other label part is switched off, in the same
    // tolerated batch as the surface, and a waterfall takes no position.
    expect(chart?.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
      showPercentage: false,
      showLegendKey: false,
      showBubbleSize: false,
    });
    expect(chart?.dataLabels.position).toBeUndefined();
    expect(chart?.dataLabels.font.name).toBe(palette.font);
    expect(chart?.axes.category.majorGridlines).toBe(false);
    expect(chart?.series[0]?.showConnectorLines).toBe(true);
    // Totals branded by position, a fall in the external colour.
    expect(chart?.series[0]?.pointColors).toEqual({
      0: palette.primary,
      1: palette.accent,
      2: palette.external,
      3: palette.primary,
    });
  });

  it("reports the gap when the deltas miss the closing total", async () => {
    seedBridge(95);
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: deltas imply 90, closing total says 95",
    );
  });

  it("refuses a bridge that is not two lines of numbers", async () => {
    helpers.select("Model!A1:A5");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "Select labels and values in two adjacent columns, or in two adjacent rows, with three or more points.",
    );

    helpers.seed("Model!D1", [
      ["Opening", 100],
      ["Price", "n/a"],
      ["Closing", 120],
    ]);
    helpers.select("Model!D1:E3");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "The second column must hold numbers only.",
    );
  });

  it("caps a bridge at a hundred points", async () => {
    helpers.select("Model!A1:B101");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "A bridge chart supports up to 100 points.",
    );
  });

  it("asks for a chart before restyling one", async () => {
    expect(await rejects(() => smt.formatSelectedChart())).toBe(
      "Select a chart first.",
    );
  });

  it("restyles the selected chart and brands every series", async () => {
    seedBridge(90);
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.chartType = "ColumnClustered";
    chart.seriesCount = 3;
    chart.title = "kept";

    await smt.formatSelectedChart();

    // A restyle never renames the chart.
    expect(chart.title).toBe("kept");
    expect(chart.borderLineStyle).toBe("None");
    expect(chart.roundedCorners).toBe(false);
    expect(chart.legend.position).toBe("Bottom");
    expect(chart.legend.visible).toBe(true);
    expect(chart.axes.value.fontName).toBe(palette.font);
    expect(chart.series.map((entry) => entry.fillColor)).toEqual([
      palette.accent,
      palette.primary,
      brand.tint(palette.primary, 0.55),
    ]);
    // Labels carry the values only, outside the end of a clustered column.
    expect(chart.dataLabels).toMatchObject({
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
      showPercentage: false,
      showLegendKey: false,
      showBubbleSize: false,
      position: "OutsideEnd",
    });
    expect(chart.dataLabels.font).toMatchObject({
      name: palette.font,
      size: 9,
    });
  });

  it("keeps a pie's legend for its categories and labels the values outside", async () => {
    seedBridge(90);
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.chartType = "Pie";
    chart.seriesCount = 1;
    chart.axes = { category: {}, value: {} };

    await smt.formatSelectedChart();

    expect(chart.axes.category).toEqual({});
    // The labels no longer name the slices, so the legend has to.
    expect(chart.legend.visible).toBe(true);
    expect(chart.dataLabels).toMatchObject({
      showValue: true,
      showPercentage: false,
      showCategoryName: false,
      position: "OutsideEnd",
      showLeaderLines: true,
    });
  });

  it("hides the legend of a lone series and places labels by chart type", async () => {
    seedBridge(90);
    await smt.insertWaterfall();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    chart.chartType = "Line";
    chart.seriesCount = 1;

    await smt.formatSelectedChart();
    expect(chart.legend.visible).toBe(false);
    expect(chart.dataLabels.position).toBe("Top");

    chart.chartType = "ColumnStacked";
    await smt.formatSelectedChart();
    expect(chart.dataLabels.position).toBe("Center");

    // Excel offers a doughnut no label position; the restyle sets none.
    chart.chartType = "Doughnut";
    chart.dataLabels.position = undefined;
    await smt.formatSelectedChart();
    expect(chart.dataLabels.position).toBeUndefined();
  });

  it("refuses to restyle on a host without the chart API", async () => {
    helpers.setSupported((_set, version) => version !== "1.9");
    expect(await rejects(() => smt.formatSelectedChart())).toBe(
      "Chart formatting needs a newer Excel build.",
    );
  });

  it("places a CAGR callout beside the series", async () => {
    helpers.seed("Model!A1", [[100, 110, 121, 133.1]]);
    helpers.select("Model!A1:D1");

    expect(await smt.addCagrLabel()).toBe("CAGR +10.0% over 3 periods");

    const shape = workbook.shapes[0];
    expect(shape).toMatchObject({
      sheetName: "Model",
      text: "CAGR +10.0%",
      width: 104,
      height: 20,
      left: 264,
      top: 0,
      fillCleared: true,
      lineVisible: false,
    });
    expect(shape?.textFrame).toEqual({
      horizontalAlignment: "Left",
      verticalAlignment: "Middle",
      font: { name: palette.font, size: 11, bold: true, color: palette.accent },
    });
  });

  it("drops the callout where Excel wants it when geometry is unavailable", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    helpers.seed("Model!A1", [[100, 121]]);
    helpers.select("Model!A1:B1");

    await smt.addCagrLabel();
    expect(workbook.shapes[0]?.left).toBeUndefined();
    expect(workbook.shapes[0]?.top).toBeUndefined();
  });

  it("refuses a callout without two numbers or a shape API", async () => {
    helpers.seed("Model!A1", [["text", 121]]);
    helpers.select("Model!A1:B1");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "The first and last cells must hold numbers.",
    );

    helpers.seed("Model!A1", [[100, 121]]);
    helpers.setSupported((_set, version) => version !== "1.9");
    expect(await rejects(() => smt.addCagrLabel())).toBe(
      "Chart labels need a newer Excel build.",
    );
  });
});

// ---------------------------------------------------------------------------

describe("tornado", () => {
  // A title over the labels, the base over the outcomes, then a header row.
  function seedDrivers(): void {
    helpers.seed("Model!A1", [["EBITDA sensitivity", 100]]);
    helpers.seed("Model!A2", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
      ["Price", 60, 140],
      ["Mix", 95, 105],
    ]);
    helpers.select("Model!A2:C5");
  }

  it("ranks the drivers into a helper block and charts it", async () => {
    seedDrivers();
    expect(await smt.insertTornado()).toBe(
      "Tornado added: 3 drivers, base 100",
    );

    // Widest swing first, both series as deltas from the base.
    expect(helpers.value("Model!D2")).toBe("Driver");
    expect(helpers.value("Model!E2")).toBe("Low");
    expect(helpers.value("Model!F2")).toBe("High");
    expect(
      ["D3", "E3", "F3"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Price", -40, 40]);
    expect(
      ["D4", "E4", "F4"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Volume", -10, 15]);
    expect(
      ["D5", "E5", "F5"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Mix", -5, 5]);

    const chart = workbook.charts[0];
    expect(chart).toMatchObject({
      chartType: "BarClustered",
      sourceAddress: "Model!D2:F5",
      seriesBy: "Columns",
      seriesCount: 2,
      title: "EBITDA sensitivity",
    });
    // Categories reversed, so the widest swing sits on top of the bars.
    expect(chart?.axes.category.reversePlotOrder).toBe(true);
    expect(chart?.series).toEqual([
      {
        fillColor: palette.external,
        overlap: 100,
        gapWidth: 40,
        pointColors: {},
      },
      {
        fillColor: palette.accent,
        overlap: 100,
        gapWidth: 40,
        pointColors: {},
      },
    ]);
  });

  it("reads a header-free table and falls back to the mean base", async () => {
    helpers.seed("Model!A1", [
      ["Volume", 80, 120],
      ["Price", 90, 110],
    ]);
    helpers.select("Model!A1:C2");

    expect(await smt.insertTornado()).toBe(
      "Tornado added: 2 drivers, base 100",
    );
    expect(
      ["D2", "E2", "F2"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Volume", -20, 20]);
    expect(
      ["D3", "E3", "F3"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Price", -10, 10]);
    expect(workbook.charts[0]?.title).toBe("Sensitivity");
  });

  // The block has to be free of values to be written at all, so what undo has
  // to put back is the formatting that stood on those cells.
  it("puts the helper block back on undo", async () => {
    helpers.setFill("Model!D2:F5", { color: "#FFEECC", pattern: "Solid" });
    seedDrivers();
    await expectExactUndo(() => smt.insertTornado());
  });

  it("leaves the bar shaping alone on an older host", async () => {
    helpers.setSupported(
      (_set, version) => version !== "1.7" && version !== "1.8",
    );
    seedDrivers();
    await smt.insertTornado();

    expect(workbook.charts[0]?.axes.category.reversePlotOrder).toBeUndefined();
    expect(workbook.charts[0]?.series[0]?.overlap).toBeUndefined();
    expect(workbook.charts[0]?.series[0]?.gapWidth).toBeUndefined();
  });

  it("refuses anything that is not a driver, low, high table", async () => {
    const shape =
      "tornado: need 3 columns (label, low, high) and at least 2 rows";
    helpers.select("Model!A1:B5");
    expect(await rejects(() => smt.insertTornado())).toBe(shape);

    // Three columns, but the header leaves a single driver behind.
    helpers.seed("Model!A1", [
      ["Driver", "Low", "High"],
      ["Volume", 90, 115],
    ]);
    helpers.select("Model!A1:C2");
    expect(await rejects(() => smt.insertTornado())).toBe(shape);

    helpers.seed("Model!A1", [
      ["Volume", 90, "n/a"],
      ["Price", 60, 140],
    ]);
    helpers.select("Model!A1:C2");
    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado: the low and high columns must hold numbers",
    );

    helpers.select("Model!A1:C101");
    expect(await rejects(() => smt.insertTornado())).toBe(
      "tornado: supports up to 100 drivers",
    );
  });
});

// ---------------------------------------------------------------------------

describe("consistent rounding", () => {
  const SHAPE_ERROR =
    "Consistent rounding: select one row or column with at least two numbers.";

  function seedColumn(): void {
    helpers.seed("Model!B2", [[33.333], [33.333], [33.334]]);
    helpers.select("Model!B2:B4");
  }

  it("writes one PLSFIX.ROUND per cell into the column beside the selection", async () => {
    seedColumn();
    expect(await smt.insertConsistentRounding()).toBe(
      "Consistent rounding: 3 cells at 0 decimals",
    );

    // Every cell points at the whole group, absolutely, and reads out its own
    // position; the source column is untouched.
    expect(
      ["C2", "C3", "C4"].map((at) => helpers.formula(`Model!${at}`)),
    ).toEqual([
      "=PLSFIX.ROUND($B$2:$B$4,1,0)",
      "=PLSFIX.ROUND($B$2:$B$4,2,0)",
      "=PLSFIX.ROUND($B$2:$B$4,3,0)",
    ]);
    expect(helpers.value("Model!B2")).toBe(33.333);
  });

  it("writes below a selected row and takes the decimals from the format", async () => {
    helpers.seed("Model!B2", [[1.005, 2.005, 3.005]]);
    helpers.setNumberFormat("Model!B2:D2", "#,##0.00");
    helpers.select("Model!B2:D2");

    expect(await smt.insertConsistentRounding()).toBe(
      "Consistent rounding: 3 cells at 2 decimals",
    );
    expect(
      ["B3", "C3", "D3"].map((at) => helpers.formula(`Model!${at}`)),
    ).toEqual([
      "=PLSFIX.ROUND($B$2:$D$2,1,2)",
      "=PLSFIX.ROUND($B$2:$D$2,2,2)",
      "=PLSFIX.ROUND($B$2:$D$2,3,2)",
    ]);
  });

  // A percentage prints two places left of what it stores, so the shares add up
  // as they are read, not as they are kept.
  it("reads a percentage format as the precision of the stored value", async () => {
    helpers.seed("Model!A1", [[0.31428], [0.31428], [0.37144]]);
    helpers.setNumberFormat("Model!A1:A3", "0.0%");
    helpers.select("Model!A1:A3");

    expect(await smt.insertConsistentRounding()).toBe(
      "Consistent rounding: 3 cells at 3 decimals",
    );
    expect(helpers.formula("Model!B1")).toBe("=PLSFIX.ROUND($A$1:$A$3,1,3)");
  });

  it("refuses to overwrite what already stands beside the selection", async () => {
    seedColumn();
    helpers.seed("Model!C3", [["base case"]]);
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding: the cells right of the selection are not empty.",
    );
    expect(helpers.formula("Model!C2")).toBe("");

    helpers.seed("Model!A1", [[1, 2]]);
    helpers.seed("Model!B2", [[7]]);
    helpers.select("Model!A1:B1");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding: the cells below the selection are not empty.",
    );
  });

  it("refuses anything that is not a line of numbers", async () => {
    helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding: select a single row or a single column, not a block.",
    );

    helpers.select("Model!A1");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      SHAPE_ERROR,
    );

    helpers.seed("Model!D1", [[10], ["n/a"], [30]]);
    helpers.select("Model!D1:D3");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      SHAPE_ERROR,
    );
  });

  it("puts the column back on undo", async () => {
    seedColumn();
    await expectExactUndo(() => smt.insertConsistentRounding());
  });
});

// ---------------------------------------------------------------------------

describe("unpivot", () => {
  function seedCrossTab(): void {
    helpers.seed("Model!A1", [
      ["", "North", "South"],
      ["Q1", 10, 20],
      ["Q2", 30, ""],
    ]);
    helpers.select("Model!A1:C3");
  }

  it("writes the long table on a new sheet and goes there", async () => {
    seedCrossTab();
    expect(await smt.unpivotSelection()).toBe("Unpivot: 3 rows on Unpivot");

    expect(
      ["A1", "B1", "C1"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Row", "Column", "Value"]);
    expect(
      ["A2", "B2", "C2"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Q1", "North", 10]);
    expect(
      ["A3", "B3", "C3"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Q1", "South", 20]);
    expect(
      ["A4", "B4", "C4"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Q2", "North", 30]);
    // The blank cell writes no line at all.
    expect(helpers.value("Unpivot!A5")).toBe("");

    expect(helpers.font("Unpivot!A1")).toMatchObject({
      name: palette.font,
      size: 10,
      bold: true,
      color: theme.titleText,
    });
    expect(helpers.fill("Unpivot!A1").color).toBe(theme.titleFill);
    expect(helpers.font("Unpivot!C4")).toMatchObject({
      name: palette.font,
      bold: false,
    });
    expect(workbook.activeSheetId).toBe(helpers.sheet("Unpivot").id);
    // The source is read, never rewritten.
    expect(helpers.value("Model!B2")).toBe(10);
  });

  it("takes the next free name when an Unpivot sheet exists", async () => {
    helpers.addSheet("unpivot");
    seedCrossTab();

    expect(await smt.unpivotSelection()).toBe("Unpivot: 3 rows on Unpivot 2");
    expect(helpers.value("Unpivot 2!A2")).toBe("Q1");
  });

  it("refuses a selection it cannot key, empty or oversized", async () => {
    helpers.select("Model!A1:B1");
    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: need a header row, a key column and one column of values",
    );

    helpers.seed("Model!A1", [
      ["", "North"],
      ["Q1", ""],
    ]);
    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: the selection holds no values",
    );

    helpers.select("Model!A1:Z500");
    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "Unpivot supports up to 5,000 selected cells at once.",
    );
  });
});

// ---------------------------------------------------------------------------

describe("contents sheet", () => {
  beforeEach(async () => {
    await boot({ sheets: ["Model", "Data", "Notes"] });
  });

  it("builds a linked contents sheet in front of the workbook", async () => {
    await smt.insertToc();

    const toc = helpers.sheet("TOC");
    expect(toc.position).toBe(0);
    expect(helpers.value("TOC!A1")).toBe("pls,fix - Contents");

    expect(helpers.value("TOC!A3")).toBe(1);
    expect(helpers.value("TOC!B3")).toBe("Model");
    expect(helpers.cell("TOC!B3").hyperlink).toEqual({
      documentReference: "'Model'!A1",
      textToDisplay: "Model",
    });
    expect(helpers.value("TOC!B5")).toBe("Notes");

    expect(helpers.font("TOC!A1")).toMatchObject({
      name: palette.font,
      size: 15,
      bold: true,
      color: theme.titleText,
    });
    expect(helpers.fill("TOC!A1").color).toBe(theme.titleFill);
    expect(helpers.rowHeight("TOC", 0)).toBe(25);
    expect(helpers.cell("TOC!A3").horizontalAlignment).toBe("Right");
    // Excel dresses a new link in its own blue; the brand colour goes on after.
    expect(helpers.font("TOC!B3")).toMatchObject({
      color: theme.linkFont,
      underline: "None",
    });
    expect(helpers.columnWidth("TOC", 0)).toBe(34);
    expect(helpers.columnWidth("TOC", 1)).toBe(240);
    expect(toc.showGridlines).toBe(false);
    expect(workbook.activeSheetId).toBe(toc.id);
  });

  it("skips hidden sheets and never lists itself", async () => {
    helpers.sheet("Notes").visibility = "Hidden";
    await smt.insertToc();

    expect(helpers.value("TOC!B3")).toBe("Model");
    expect(helpers.value("TOC!B4")).toBe("Data");
    expect(helpers.value("TOC!B5")).toBe("");
  });

  it("rewrites its own sheet without leaving stale rows behind", async () => {
    await smt.insertToc();
    expect(helpers.value("TOC!B5")).toBe("Notes");

    helpers.deleteSheet("Notes");
    await smt.insertToc();

    expect(helpers.value("TOC!A1")).toBe("pls,fix - Contents");
    expect(helpers.value("TOC!B4")).toBe("Data");
    expect(helpers.value("TOC!B5")).toBe("");
    expect(helpers.cell("TOC!B5").hyperlink).toBeNull();
  });

  it("refuses to overwrite a TOC sheet it did not write", async () => {
    helpers.addSheet("TOC");
    helpers.seed("TOC!A1", [["My own contents"]]);

    expect(await rejects(() => smt.insertToc())).toBe(
      "A sheet named TOC already exists and is not ours.",
    );
    expect(helpers.value("TOC!A1")).toBe("My own contents");
  });
});

// ---------------------------------------------------------------------------

describe("sheet explorer", () => {
  it("lists every sheet and marks the active one", async () => {
    await boot({ sheets: ["Model", "Data", "Notes"] });
    helpers.sheet("Notes").visibility = "VeryHidden";
    await smt.activateSheet("Data");

    expect(await smt.listSheets()).toEqual([
      { name: "Model", visibility: "Visible", active: false },
      { name: "Data", visibility: "Visible", active: true },
      { name: "Notes", visibility: "VeryHidden", active: false },
    ]);
  });

  it("hides and shows a sheet", async () => {
    await smt.setSheetVisibility("Data", false);
    expect(helpers.sheet("Data").visibility).toBe("Hidden");

    await smt.setSheetVisibility("Data", true);
    expect(helpers.sheet("Data").visibility).toBe("Visible");
  });

  it("leaves a very hidden sheet alone", async () => {
    helpers.sheet("Data").visibility = "VeryHidden";
    expect(await rejects(() => smt.setSheetVisibility("Data", true))).toBe(
      "Data is very hidden and can only be shown in VBA.",
    );
    expect(helpers.sheet("Data").visibility).toBe("VeryHidden");
  });

  it("refuses to hide the last visible sheet", async () => {
    await boot({ sheets: ["Only"] });
    expect(await rejects(() => smt.setSheetVisibility("Only", false))).toBe(
      "A workbook needs at least one visible sheet.",
    );
    expect(helpers.sheet("Only").visibility).toBe("Visible");
  });

  it("activates a sheet by name", async () => {
    await smt.activateSheet("Data");
    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
  });
});

// ---------------------------------------------------------------------------

describe("broken names", () => {
  beforeEach(() => {
    helpers.addName("Revenue", "=Model!$A$1");
    helpers.addName("Costs", "=Model!#REF!");
    helpers.addName("Margin", "=#REF!$B$2");
  });

  it("finds every name pointing at a deleted range", async () => {
    expect(await smt.listBrokenNames()).toEqual(["Costs", "Margin"]);
  });

  it("deletes only the broken ones", async () => {
    expect(await smt.deleteBrokenNames()).toBe(2);
    expect(workbook.names.map((entry) => entry.name)).toEqual(["Revenue"]);
    expect(await smt.listBrokenNames()).toEqual([]);
  });
});

// ---------------------------------------------------------------------------

describe("smart track", () => {
  beforeEach(() => {
    helpers.select("Model!B2");
  });

  it("lists direct precedents area by area", async () => {
    helpers.setPrecedents("Model!B2", [
      { address: "Model!A1:A3", cellCount: 3 },
      { address: "Data!C1", cellCount: 1 },
    ]);

    expect(await smt.traceActiveCell("precedents")).toEqual({
      origin: "Model!B2",
      areas: [
        { sheet: "Model", address: "A1:A3", cellCount: 3 },
        { sheet: "Data", address: "C1", cellCount: 1 },
      ],
    });
  });

  it("lists direct dependents", async () => {
    helpers.setDependents("Model!B2", [{ address: "Model!D9", cellCount: 1 }]);

    expect(await smt.traceActiveCell("dependents")).toEqual({
      origin: "Model!B2",
      areas: [{ sheet: "Model", address: "D9", cellCount: 1 }],
    });
  });

  it("reads Excel's ItemNotFound as an empty result", async () => {
    helpers.setPrecedents("Model!B2", "itemNotFound");

    expect(await smt.traceActiveCell("precedents")).toEqual({
      origin: "Model!B2",
      areas: [],
    });
  });

  it("says so when the host is too old to trace", async () => {
    helpers.setSupported((_set, version) => version !== "1.12");
    expect(await rejects(() => smt.traceActiveCell("precedents"))).toBe(
      "Tracing needs a newer Excel build.",
    );

    helpers.setSupported((_set, version) => version !== "1.13");
    expect(await rejects(() => smt.traceActiveCell("dependents"))).toBe(
      "Tracing needs a newer Excel build.",
    );
  });

  it("jumps to an area on another sheet", async () => {
    await smt.selectArea({ sheet: "Data", address: "B2:C3" });

    expect(workbook.activeSheetId).toBe(helpers.sheet("Data").id);
    expect(workbook.selection.rect).toEqual({
      row: 1,
      col: 1,
      rowCount: 2,
      colCount: 2,
    });
  });

  it("jumps within the active sheet when the area names none", async () => {
    await smt.selectArea({ sheet: "", address: "D4" });
    expect(workbook.selection).toEqual({
      sheetId: helpers.sheet("Model").id,
      rect: { row: 3, col: 3, rowCount: 1, colCount: 1 },
    });
  });

  it("selects every precedent on the first area's sheet at once", async () => {
    await smt.selectAreas([
      { sheet: "Model", address: "A1", cellCount: 1 },
      { sheet: "Model", address: "C3", cellCount: 1 },
    ]);

    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
    expect(workbook.selectionAreas).toEqual([
      {
        sheetId: helpers.sheet("Model").id,
        rect: { row: 0, col: 0, rowCount: 1, colCount: 1 },
      },
      {
        sheetId: helpers.sheet("Model").id,
        rect: { row: 2, col: 2, rowCount: 1, colCount: 1 },
      },
    ]);
  });

  it("leaves a cross-sheet precedent out of the jump, for the chips to list", async () => {
    await smt.selectAreas([
      { sheet: "Model", address: "A1", cellCount: 1 },
      { sheet: "Data", address: "B2", cellCount: 1 },
    ]);

    // Only the first area's sheet was selected; Data!B2 was left alone.
    expect(workbook.activeSheetId).toBe(helpers.sheet("Model").id);
    expect(workbook.selectionAreas).toEqual([
      {
        sheetId: helpers.sheet("Model").id,
        rect: { row: 0, col: 0, rowCount: 1, colCount: 1 },
      },
    ]);
  });

  it("does nothing when there is nothing to select", async () => {
    const before = workbook.activeSheetId;
    await smt.selectAreas([]);
    expect(workbook.activeSheetId).toBe(before);
  });
});

// ---------------------------------------------------------------------------

describe("autocolor on edit", () => {
  it("registers once however often it is switched on", async () => {
    await smt.setAutocolorOnEdit(true);
    expect(helpers.changeHandlerCount()).toBe(1);

    await smt.setAutocolorOnEdit(true);
    expect(helpers.changeHandlerCount()).toBe(1);
  });

  it("colours the changed range on whichever sheet it happened", async () => {
    helpers.seed("Model!A1", [[42]]);
    helpers.seed("Data!A1", [
      [
        42,
        { formula: "=A1+C1", value: 84 },
        { formula: "=A1*1.05", value: 44 },
      ],
    ]);
    helpers.select("Model!A1");
    await smt.setAutocolorOnEdit(true);

    // The event address is sheet-qualified; only event.getRange resolves it.
    await helpers.fireChanged("Data", "A1:C1");

    expect(helpers.font("Data!A1").color).toBe(theme.inputFont);
    expect(helpers.font("Data!B1").color).toBe(theme.formulaFont);
    expect(helpers.font("Data!C1").color).toBe(theme.partialFont);
    expect(helpers.font("Model!A1").color).toBe("#000000");
  });

  it("skips a change too big to colour on the keystroke", async () => {
    helpers.seed("Data!A1", [[42]]);
    await smt.setAutocolorOnEdit(true);

    await helpers.fireChanged("Data", "A1:A501");
    expect(helpers.font("Data!A1").color).toBe("#000000");

    await helpers.fireChanged("Data", "A1:A500");
    expect(helpers.font("Data!A1").color).toBe(theme.inputFont);
  });

  it("stops colouring once it is switched off", async () => {
    helpers.seed("Data!A1", [[42]]);
    await smt.setAutocolorOnEdit(true);
    await smt.setAutocolorOnEdit(false);
    expect(helpers.changeHandlerCount()).toBe(0);

    await helpers.fireChanged("Data", "A1");
    expect(helpers.font("Data!A1").color).toBe("#000000");
  });

  it("leaves no phantom registration when the sync fails", async () => {
    helpers.failNextSync();
    await expect(smt.setAutocolorOnEdit(true)).rejects.toThrow();
    expect(helpers.changeHandlerCount()).toBe(0);

    // The queue survives the failure, so a retry still registers exactly one.
    await smt.setAutocolorOnEdit(true);
    expect(helpers.changeHandlerCount()).toBe(1);
  });

  it("keeps the handle when the removal sync fails, so nothing is orphaned", async () => {
    await smt.setAutocolorOnEdit(true);
    helpers.failNextSync();
    await expect(smt.setAutocolorOnEdit(false)).rejects.toThrow();
    // Removal never synced: still registered, and the handle is retained.
    expect(helpers.changeHandlerCount()).toBe(1);

    // Re-enable must not stack a second registration on top.
    await smt.setAutocolorOnEdit(true);
    expect(helpers.changeHandlerCount()).toBe(1);

    // The retained handle makes a retried disable land cleanly.
    await smt.setAutocolorOnEdit(false);
    expect(helpers.changeHandlerCount()).toBe(0);
  });

  it("registers a ribbon command through the actions registry", () => {
    const actions = helpers.actions();
    const host = globalThis as unknown as {
      Office: { actions: { associate: (id: string, fn: () => void) => void } };
    };
    host.Office.actions.associate("PLSFIX_PROBE", () => undefined);
    expect(actions.has("PLSFIX_PROBE")).toBe(true);
  });
});

// ---------------------------------------------------------------------------

describe("resolution by sheet id", () => {
  it("pastes after the source sheet was renamed", async () => {
    helpers.seed("Data!A1", [[1, 2]]);
    helpers.select("Data!A1:B1");
    await smt.markCopySource();

    helpers.sheet("Data").name = "Renamed";
    helpers.select("Model!D1");
    await smt.pasteSpecial("values");

    expect(helpers.value("Model!E1")).toBe(2);
  });

  it("undoes after the sheet it ran on was renamed", async () => {
    helpers.seed("Data!A1", [[1, 2]]);
    const before = helpers.cellMap("Data");
    helpers.select("Data!A1:B1");
    await smt.applyPreset("header");

    helpers.sheet("Data").name = "Renamed";
    expect(await smt.undoLastAction()).toBe(
      "Undone: Data!A1:B1. Nothing more to undo.",
    );
    expect(helpers.cellMap("Renamed")).toEqual(before);
  });

  it("puts the old fills back before snapshotting an overlapping block", async () => {
    helpers.seed("Model!A1", [
      [{ formula: "=B1", r1c1: "=RC[1]", value: 1 }, 1],
      [{ formula: "=B2", r1c1: "=RC[1]", value: 2 }, 2],
      [{ formula: "=B3", r1c1: "=RC[1]", value: 3 }, 3],
    ]);
    const before = helpers.cellMap("Model");

    helpers.select("Model!A1:B2");
    await smt.toggleAuditOverlay();
    // Overlaps the painted block: the old fills go back before the new snapshot,
    // or the overlay would capture its own stripes as the modeller's formatting.
    helpers.select("Model!A2:B3");
    await smt.toggleAuditOverlay();
    await smt.toggleAuditOverlay();

    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("skips a saved overlay whose sheet is gone", async () => {
    helpers.seed("Data!A1", [
      [{ formula: "=B1", r1c1: "=RC[1]", value: 1 }, 1],
      [{ formula: "=B2", r1c1: "=RC[1]", value: 2 }, 2],
    ]);
    helpers.select("Data!A1:B2");
    await smt.toggleAuditOverlay();

    helpers.deleteSheet("Data");
    await reopen();

    expect(await smt.restorePersistedOverlay()).toBe(false);
    expect(workbook.settings.get("smtAuditOverlay")).toBe("");
  });
});

// ---------------------------------------------------------------------------

// The instrument itself: without these, a green suite would only prove that
// nothing above happens to read an unloaded property in the fake's own way.
describe("strict load semantics", () => {
  interface Probe {
    load: (properties?: string) => void;
    address: string;
    values: unknown;
    format: { fill: { color: string } };
    getCell: (row: number, column: number) => Probe;
  }

  async function inRun(
    body: (context: Excel.RequestContext, range: Probe) => Promise<void>,
  ): Promise<void> {
    const host = globalThis as unknown as {
      Excel: { run: (cb: (context: never) => unknown) => Promise<unknown> };
    };
    await host.Excel.run(async (context: never) => {
      const workbook = (
        context as unknown as {
          workbook: { getSelectedRange: () => Probe };
        }
      ).workbook;
      await body(context as Excel.RequestContext, workbook.getSelectedRange());
    });
  }

  async function throws(run: () => void): Promise<string> {
    try {
      run();
    } catch (error) {
      return (error as Error).message;
    }
    return "nothing was thrown";
  }

  beforeEach(() => {
    helpers.seed("Model!A1", [[1, 2]]);
    helpers.select("Model!A1:B1");
  });

  it("refuses a scalar read with no load behind it", async () => {
    await inRun(async (_context, range) => {
      expect(await throws(() => range.address)).toBe(
        "The property 'address' is not available. Before reading the " +
          "property's value, call the load method on the containing object " +
          'and call "context.sync()".',
      );
    });
  });

  it("refuses a load that has not been synced yet", async () => {
    await inRun(async (_context, range) => {
      range.load("address");
      expect(await throws(() => range.address)).toContain("is not available");
    });
  });

  it("serves a property once it is loaded and synced", async () => {
    await inRun(async (context, range) => {
      range.load("address,values");
      await context.sync();
      expect(range.address).toBe("Model!A1:B1");
      expect(range.values).toEqual([[1, 2]]);
    });
  });

  it("unlocks one property at a time, not the whole object", async () => {
    await inRun(async (context, range) => {
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[1, 2]]);
      expect(await throws(() => range.address)).toContain("is not available");
    });
  });

  it("keys the unlock to the proxy, not to the cells behind it", async () => {
    await inRun(async (context, range) => {
      range.load("values");
      await context.sync();
      // A fresh proxy over the same cells starts with nothing loaded, which is
      // what makes a second Excel.run a clean slate.
      expect(await throws(() => range.getCell(0, 0).values)).toContain(
        "is not available",
      );
    });
  });

  it("reaches nested paths and leaves navigation alone", async () => {
    await inRun(async (context, range) => {
      expect(await throws(() => range.format.fill.color)).toContain(
        "The property 'color' is not available",
      );
      range.load("format/fill/color");
      await context.sync();
      expect(range.format.fill.color).toBe("#FFFFFF");
    });
  });

  it("commits nothing when the sync fails", async () => {
    await inRun(async (context, range) => {
      range.load("address");
      helpers.failNextSync();
      await expect(context.sync()).rejects.toThrow();
      expect(await throws(() => range.address)).toContain("is not available");
    });
  });
});

// ---------------------------------------------------------------------------

describe("separators", () => {
  it("reads the decimal and thousands separators Excel is set to", async () => {
    await boot({ separators: { decimal: ",", thousands: " " } });
    expect(await smt.readSeparators()).toEqual({
      decimal: ",",
      thousands: " ",
    });
  });

  it("is null on a host below ExcelApi 1.11", async () => {
    await boot({
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.11"),
    });
    expect(await smt.readSeparators()).toBeNull();
  });
});
