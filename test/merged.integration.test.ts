// Selections a real model hands the pane: a merged title band over the numbers,
// and blocks with nothing in them yet. Neither may stop a tool. A merged block
// takes its value in the top-left cell only, and Excel refuses a value write
// that covers part of one, which is the one case worth a message of its own.

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
  const host = installFakeHost({ sheets: ["Model"] });
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

// A merged heading over two columns, then two rows of numbers under it.
function seedMergedBlock(): void {
  helpers.seed("Model!A1", [["Revenue bridge"]]);
  helpers.merge("Model!A1:B1");
  helpers.seed("Model!A2", [
    [10, 20],
    [30, 40],
  ]);
  helpers.select("Model!A1:B3");
}

beforeEach(async () => {
  await boot();
});

describe("a selection holding a merged block", () => {
  it("flips the sign of the numbers under it", async () => {
    seedMergedBlock();
    await smt.applySignFlip();

    expect(helpers.value("Model!A2")).toBe(-10);
    expect(helpers.value("Model!B3")).toBe(-40);
    expect(helpers.value("Model!A1")).toBe("Revenue bridge");
  });

  it("takes a number format", async () => {
    seedMergedBlock();
    await smt.applyNumberFormat("whole");
    expect(helpers.numberFormat("Model!B3")).toBe("#,##0;[Red](#,##0);-");
  });

  it("takes a preset and a fill cycle", async () => {
    seedMergedBlock();
    await smt.applyPreset("input");
    await smt.applyFillCycle();
    expect(helpers.font("Model!B3").size).toBe(10);
    expect(helpers.fill("Model!A1").pattern).not.toBe("");
  });

  it("steps its decimals", async () => {
    seedMergedBlock();
    helpers.setNumberFormat("Model!A2:B3", "#,##0.0");
    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!B3")).toBe("#,##0.00");
  });

  it("guards its formulas", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:B1");
    helpers.seed("Model!A2", [[{ formula: "=X1/0", value: 0 }]]);
    helpers.select("Model!A1:B2");

    await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A2")).toBe("=IFERROR(X1/0,0)");
  });

  it("says what to do when the selection cuts one", async () => {
    seedMergedBlock();
    // Column A only: the merged A1:B1 band is half in, half out.
    helpers.select("Model!A1:A3");

    expect(await rejects(() => smt.applySignFlip())).toBe(
      "Sign flip: Excel refused this write. Select whole merged cells, not part of one.",
    );
    expect(helpers.value("Model!A2")).toBe(10);
  });
});

describe("a selection of empty cells", () => {
  beforeEach(() => {
    helpers.select("Model!C3:E6");
  });

  it("takes a number format", async () => {
    await smt.applyNumberFormat("percent");
    expect(helpers.numberFormat("Model!D4")).toBe("0.0%;[Red](0.0%);-");
  });

  it("takes the cycles", async () => {
    await smt.applyNumberCycle("currency");
    await smt.applyFillCycle();
    await smt.applyFontColorCycle();
    await smt.applyRowStyleCycle("item");
    await smt.applyBorderCycle();
    expect(helpers.numberFormat("Model!C3")).not.toBe("General");
  });

  it("takes a sign flip, a scale and the IFERROR guard", async () => {
    await smt.applySignFlip();
    await smt.scaleSelection(1000);
    await smt.toggleIfErrorGuard();
    expect(helpers.value("Model!D4")).toBe("");
  });

  it("takes the decimal steppers", async () => {
    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!D4")).toBe("0.0");

    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!D4")).toBe("0");
  });

  it("takes a preset and gives it back on undo", async () => {
    const before = helpers.cellMap("Model");
    await smt.applyPreset("header");
    await smt.undoLastAction();
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});
