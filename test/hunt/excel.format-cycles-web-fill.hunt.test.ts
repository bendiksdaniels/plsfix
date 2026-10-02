// format-cycles.ts keeps its own private fill reader instead of routing
// through fill-store.ts's readFill. Excel for the web answers a solid fill
// with a null pattern and an unfilled cell with colour "" (fill-store.ts,
// rig 27.09), so the local reader's `pattern === Excel.FillPattern.none`
// check never fires for a genuinely unfilled cell on the web: it falls
// through to `fill.color.toUpperCase()`, reading "" instead of the CLEAR_FILL
// sentinel. The row-style cycle (readCellStyle) then never recognises the
// Plain step as matching any variant, so `matchStyleIndex` misses and the
// cycle always restarts at index 0 instead of advancing from wherever Plain
// actually sits.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import type * as SettingsModule from "../../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let theme: ReturnType<(typeof SettingsModule)["deriveTheme"]>;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"], web: true }).helpers;
  smt = await import("../../src/excel");
  const brand: typeof SettingsModule = await import("../../src/settings");
  theme = brand.deriveTheme(brand.DEFAULT_SETTINGS);
});

describe("row style cycle on Excel for the web", () => {
  it("steps the item cycle exactly like every other host", async () => {
    helpers.select("Model!A1:A2");

    // First press: whatever the starting state, it lands the cycle on Plain.
    await smt.applyRowStyleCycle("item");
    expect(helpers.fill("Model!A1").pattern).toBe("None");

    // Second press: from a REAL Plain cell, the cycle must advance to the
    // next look (header tint), exactly as it does on every other host
    // (test/host.integration.test.ts "steps the item cycle"). The buggy
    // reader instead re-matches nothing, resets to index 0 and writes
    // Plain again - "None" a second time instead of the header colour.
    await smt.applyRowStyleCycle("item");
    expect(helpers.fill("Model!A1").pattern).toBe("Solid");
    expect(helpers.fill("Model!A1").color).toBe(theme.headerFill);
  });

  it("never gets stuck: three presses visit three distinct looks", async () => {
    helpers.select("Model!A1");
    const receipts = [
      await smt.applyRowStyleCycle("item"),
      await smt.applyRowStyleCycle("item"),
      await smt.applyRowStyleCycle("item"),
    ];
    expect(new Set(receipts).size).toBe(3);
    expect(receipts).toEqual([
      "Row style: Plain",
      "Row style: Header tint",
      "Row style: Result tint",
    ]);
  });

  it("advances the result cycle's Plain middle step forward, not backward", async () => {
    helpers.select("Model!A1");
    // Press 1: unrecognised starting format -> index 0 (Filled/resultFill).
    const first = await smt.applyRowStyleCycle("result");
    expect(first).toBe("Row style: Filled");
    // Press 2: matched at index 0 for real -> advances to Plain (index 1).
    const second = await smt.applyRowStyleCycle("result");
    expect(second).toBe("Row style: Double rule");
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    // Press 3: the cell is now GENUINELY Plain (fill cleared, font colour and
    // bold set exactly as the Plain variant demands). It must advance to
    // Accent block (index 2), not fall back to Filled (index 0) because the
    // web read of "no fill" was mistaken for an unrecognised colour.
    const third = await smt.applyRowStyleCycle("result");
    expect(third).toBe("Row style: Accent block");
  });
});
