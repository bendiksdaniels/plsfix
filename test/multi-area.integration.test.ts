// A ctrl-clicked selection. office.js refuses to serve one through
// getSelectedRange at all, so every formatting and editing action reopens the
// areas from Excel.RangeAreas and treats them as one job: the state is read
// from the first area, and the write lands in all of them.

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
let brand: typeof SettingsModule;
let theme: ReturnType<(typeof SettingsModule)["deriveTheme"]>;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
  theme = brand.deriveTheme(brand.DEFAULT_SETTINGS);
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// Two blocks a modeller would ctrl-click: a cost column and a revenue column.
function selectTwoBlocks(): void {
  helpers.seed("Model!A1", [[1], [2], [3]]);
  helpers.seed("Model!C1", [[10], [20], [30]]);
  helpers.selectAreas(["Model!A1:A3", "Model!C1:C3"]);
}

beforeEach(async () => {
  await boot();
});

describe("formatting every area", () => {
  it("writes a number format into both blocks", async () => {
    selectTwoBlocks();
    await smt.applyNumberFormat("whole");

    expect(helpers.numberFormat("Model!A3")).toBe("#,##0;[Red](#,##0);-");
    expect(helpers.numberFormat("Model!C3")).toBe("#,##0;[Red](#,##0);-");
  });

  it("paints a preset over both blocks", async () => {
    selectTwoBlocks();
    await smt.applyPreset("input");

    expect(helpers.font("Model!A2").color).toBe(theme.inputFont);
    expect(helpers.font("Model!C2").color).toBe(theme.inputFont);
  });

  it("clears formats in both blocks", async () => {
    selectTwoBlocks();
    helpers.setFill("Model!A1:A3", { color: "#FFEECC", pattern: "Solid" });
    helpers.setFill("Model!C1:C3", { color: "#FFEECC", pattern: "Solid" });

    await smt.clearFormats();

    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.fill("Model!C1").pattern).toBe("None");
  });

  it("steps one cycle for both blocks, reading the first cell", async () => {
    selectTwoBlocks();
    await smt.applyFillCycle();

    const first = helpers.fill("Model!A1").color;
    expect(helpers.fill("Model!C3").color).toBe(first);
    expect(first).not.toBe("#FFFFFF");
  });

  it("steps the font colour cycle in both blocks", async () => {
    selectTwoBlocks();
    await smt.applyFontColorCycle();

    expect(helpers.font("Model!C1").color).toBe(helpers.font("Model!A1").color);
  });

  it("steps the number cycle in both blocks", async () => {
    selectTwoBlocks();
    await smt.applyNumberCycle("percent");

    expect(helpers.numberFormat("Model!C2")).toBe(
      helpers.numberFormat("Model!A2"),
    );
    expect(helpers.numberFormat("Model!C2")).not.toBe("General");
  });

  it("steps the row style cycle in both blocks", async () => {
    selectTwoBlocks();
    await smt.applyRowStyleCycle("title");

    expect(helpers.font("Model!C3").bold).toBe(helpers.font("Model!A3").bold);
    expect(helpers.fill("Model!C1").color).toBe(helpers.fill("Model!A1").color);
  });

  it("steps the border cycle on both blocks' own edges", async () => {
    selectTwoBlocks();
    await smt.applyBorderCycle();

    expect(helpers.border("Model!C3", "bottom").style).toBe(
      helpers.border("Model!A3", "bottom").style,
    );
    expect(helpers.border("Model!C3", "bottom").style).not.toBe("None");
  });
});

describe("editing every area", () => {
  it("flips the sign in both blocks", async () => {
    selectTwoBlocks();
    await smt.applySignFlip();

    expect(helpers.value("Model!A1")).toBe(-1);
    expect(helpers.value("Model!C3")).toBe(-30);
  });

  it("scales both blocks", async () => {
    selectTwoBlocks();
    await smt.scaleSelection(1000);
    expect(helpers.value("Model!A1")).toBe(1000);
    expect(helpers.value("Model!C3")).toBe(30000);

    await smt.scaleSelection(0.001);
    expect(helpers.value("Model!A1")).toBe(1);
    expect(helpers.value("Model!C3")).toBe(30);
  });

  it("guards formulas in both blocks", async () => {
    helpers.seed("Model!A1", [[{ formula: "=B1/0", value: 0 }]]);
    helpers.seed("Model!C1", [[{ formula: "=D1/0", value: 0 }]]);
    helpers.selectAreas(["Model!A1", "Model!C1"]);

    await smt.toggleIfErrorGuard();

    expect(helpers.formula("Model!A1")).toBe("=IFERROR(B1/0,0)");
    expect(helpers.formula("Model!C1")).toBe("=IFERROR(D1/0,0)");
  });

  it("steps decimals in both blocks", async () => {
    selectTwoBlocks();
    helpers.setNumberFormat("Model!A1:A3", "#,##0.0");
    helpers.setNumberFormat("Model!C1:C3", "#,##0.0");

    await smt.applyDecimalStep(1);

    expect(helpers.numberFormat("Model!A2")).toBe("#,##0.00");
    expect(helpers.numberFormat("Model!C2")).toBe("#,##0.00");
  });

  it("pastes into every area", async () => {
    helpers.seed("Model!E1", [[7]]);
    helpers.select("Model!E1");
    await smt.markCopySource();

    selectTwoBlocks();
    await smt.pasteSpecial("values");

    expect(helpers.value("Model!A1")).toBe(7);
    expect(helpers.value("Model!C1")).toBe(7);
  });

  it("pastes formulas verbatim into every area", async () => {
    helpers.seed("Model!E1", [[{ formula: "=Z9*2", value: 4 }]]);
    helpers.select("Model!E1");
    await smt.markCopySource();

    selectTwoBlocks();
    await smt.pastePreserveFormulas();

    expect(helpers.formula("Model!A1")).toBe("=Z9*2");
    expect(helpers.formula("Model!C1")).toBe("=Z9*2");
  });
});

describe("undo over every area", () => {
  it("names every area and puts them all back", async () => {
    selectTwoBlocks();
    const before = helpers.cellMap("Model");

    await smt.applySignFlip();
    expect(smt.undoTarget()).toBe("Model!A1:A3, Model!C1:C3");
    expect(helpers.cellMap("Model")).not.toEqual(before);

    await smt.undoLastAction();
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});

describe("the caps and the older hosts", () => {
  it("counts every area against the selection cap", async () => {
    helpers.selectAreas(["Model!A1:A3000", "Model!C1:C3000"]);
    expect(await rejects(() => smt.applySignFlip())).toBe(
      "Sign flip supports up to 5,000 selected cells at once.",
    );
  });

  it("acts on a single block without RangeAreas", async () => {
    await boot();
    helpers.setSupported((_set, version) => version !== "1.9");
    helpers.seed("Model!A1", [[1], [2], [3]]);
    helpers.select("Model!A1:A3");

    await smt.applySignFlip();

    expect(helpers.value("Model!A1")).toBe(-1);
    expect(workbook.sheets.length).toBe(2);
  });

  it("says why a ctrl-clicked selection is refused on an older host", async () => {
    await boot();
    helpers.setSupported((_set, version) => version !== "1.9");
    selectTwoBlocks();

    expect(await rejects(() => smt.applySignFlip())).toBe(
      "Sign flip: this Excel build can only act on one selected block.",
    );
    expect(helpers.value("Model!A1")).toBe(1);
  });
});
