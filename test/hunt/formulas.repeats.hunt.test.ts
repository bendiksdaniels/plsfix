// Pass-1 hunt: the formula-editing actions in src/excel/formulas.ts pressed
// three times in a row, then undone to the exact original; decimal stepping
// against a host that rewrites the format on read-back, and the IFERROR guard
// over a formula guarded by hand with a foreign fallback.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  smt = await import("../../src/excel");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// Presses `run` `times` times, then undoes it exactly that many times: the
// cell map (a whole-sheet snapshot) must match what it was before the first
// press, byte for byte, no matter how many times the action stacked.
async function pressAndFullyUndo(
  run: () => Promise<unknown>,
  times: number,
  sheetName = "Model",
): Promise<void> {
  const before = helpers.cellMap(sheetName);
  for (let pass = 0; pass < times; pass += 1) await run();
  expect(helpers.cellMap(sheetName)).not.toEqual(before);
  for (let pass = 0; pass < times; pass += 1) await smt.undoLastAction();
  expect(helpers.cellMap(sheetName)).toEqual(before);
}

beforeEach(async () => {
  await boot();
});

describe("decimal stepping against a host that rewrites the currency symbol", () => {
  it("keeps counting decimals correctly across three presses", async () => {
    await boot({ rewriteCurrencyFormats: true });
    helpers.setNumberFormat("Model!A1:B1", "€ #,##0.0;[Red](€ #,##0.0);-");
    helpers.select("Model!A1:B1");

    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!A1")).toContain("[$€-x-fake]");
    expect(helpers.numberFormat("Model!A1")).toBe(
      "[$€-x-fake] #,##0.00;[Red]([$€-x-fake] #,##0.00);-",
    );

    await smt.applyDecimalStep(1);
    expect(helpers.numberFormat("Model!A1")).toBe(
      "[$€-x-fake] #,##0.000;[Red]([$€-x-fake] #,##0.000);-",
    );

    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!A1")).toBe(
      "[$€-x-fake] #,##0.00;[Red]([$€-x-fake] #,##0.00);-",
    );
  });

  it("bottoms out at zero decimals over three decreases without going negative", async () => {
    helpers.setNumberFormat("Model!A1", "#,##0.00;[Red](#,##0.00);-");
    helpers.select("Model!A1");

    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0.0;[Red](#,##0.0);-");
    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0;[Red](#,##0);-");
    await smt.applyDecimalStep(-1);
    expect(helpers.numberFormat("Model!A1")).toBe("#,##0;[Red](#,##0);-");
  });

  it("presses increase three times in a row and fully undoes it", async () => {
    helpers.setNumberFormat("Model!A1", "General");
    helpers.select("Model!A1");
    await pressAndFullyUndo(() => smt.applyDecimalStep(1), 3);
  });
});

describe("sign flip three times in a row", () => {
  it("ends negative on the third press and fully undoes back to the start", async () => {
    helpers.seed("Model!A1", [[7, { formula: "=B9+C9", value: 3 }]]);
    helpers.select("Model!A1:B1");

    await smt.applySignFlip();
    expect(helpers.value("Model!A1")).toBe(-7);
    await smt.applySignFlip();
    expect(helpers.value("Model!A1")).toBe(7);
    await smt.applySignFlip();
    expect(helpers.value("Model!A1")).toBe(-7);
    expect(helpers.formula("Model!B1")).toBe("=-(B9+C9)");

    await smt.undoLastAction();
    await smt.undoLastAction();
    await smt.undoLastAction();
    expect(helpers.value("Model!A1")).toBe(7);
    expect(helpers.formula("Model!B1")).toBe("=B9+C9");
  });
});

describe("the IFERROR guard against a formula guarded by hand", () => {
  it("strips a foreign fallback instead of nesting a second IFERROR", async () => {
    helpers.seed("Model!A1", [
      [{ formula: '=IFERROR(A9/B9,"n/a")', value: 1 }],
    ]);
    helpers.select("Model!A1");

    const receipt = await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A1")).toBe("=A9/B9");
    expect(receipt).toContain("stripped from 1 formula");
    expect(receipt).not.toContain("added");
  });

  it("cycles add/strip/add over three presses without deepening the nest", async () => {
    helpers.seed("Model!A1", [[{ formula: "=A9/B9", value: 1 }]]);
    helpers.select("Model!A1");

    await smt.toggleIfErrorGuard();
    const afterFirst = helpers.formula("Model!A1");
    expect(afterFirst).toBe("=IFERROR(A9/B9,0)");

    await smt.toggleIfErrorGuard();
    expect(helpers.formula("Model!A1")).toBe("=A9/B9");

    await smt.toggleIfErrorGuard();
    // The third press (add again) must match the first, not a deeper nest.
    expect(helpers.formula("Model!A1")).toBe(afterFirst);

    await smt.undoLastAction();
    await smt.undoLastAction();
    await smt.undoLastAction();
    expect(helpers.formula("Model!A1")).toBe("=A9/B9");
  });
});

describe("unit scaling x1000 applied twice, then reversed", () => {
  it("returns the exact original number after two ups and two downs", async () => {
    helpers.seed("Model!A1", [[5]]);
    helpers.select("Model!A1");

    await smt.scaleSelection(1000);
    expect(helpers.value("Model!A1")).toBe(5000);
    await smt.scaleSelection(1000);
    expect(helpers.value("Model!A1")).toBe(5_000_000);

    await smt.scaleSelection(0.001);
    expect(helpers.value("Model!A1")).toBe(5000);
    await smt.scaleSelection(0.001);
    expect(helpers.value("Model!A1")).toBe(5);
  });

  it("nests each of the four scales into the formula instead of simplifying it", async () => {
    helpers.seed("Model!B1", [[{ formula: "=B9", value: 2 }]]);
    helpers.select("Model!B1");

    await smt.scaleSelection(1000);
    await smt.scaleSelection(1000);
    expect(helpers.formula("Model!B1")).toBe("=((B9)*1000)*1000");

    await smt.scaleSelection(0.001);
    await smt.scaleSelection(0.001);
    // scaleCells writes the operator as text ("/1000"), not the factor
    // 0.001 itself, so each down-scale nests a division, not a multiply.
    expect(helpers.formula("Model!B1")).toBe("=((((B9)*1000)*1000)/1000)/1000");
  });
});

describe("consistent rounding pressed three times in a row", () => {
  it("refuses the second and third press identically, then undo restores the source", async () => {
    helpers.seed("Model!B2", [[33.333], [33.333], [33.334]]);
    helpers.select("Model!B2:B4");
    const before = helpers.cellMap("Model");

    await smt.insertConsistentRounding();
    const afterFirst = helpers.cellMap("Model");
    expect(afterFirst).not.toEqual(before);

    // Re-selecting the same source: the destination column is now occupied
    // by real PLSFIX.ROUND formulas, so this must refuse, not double-write.
    helpers.select("Model!B2:B4");
    const secondRefusal = await rejects(() => smt.insertConsistentRounding());
    const thirdRefusal = await rejects(() => smt.insertConsistentRounding());
    expect(secondRefusal).toBe(thirdRefusal);
    expect(secondRefusal).toBe(
      "Consistent rounding: the cells right of the selection are not empty.",
    );
    // Nothing changed on the refused presses.
    expect(helpers.cellMap("Model")).toEqual(afterFirst);

    await smt.undoLastAction();
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});

describe("CAGR pressed three times in a row on the same series", () => {
  it("overwrites the same cell identically each time and fully undoes", async () => {
    helpers.seed("Model!A2", [[100, 110, 121, 133.1]]);
    const before = helpers.cellMap("Model");

    helpers.select("Model!A2:D2");
    await smt.insertCagr();
    const afterFirst = helpers.formula("Model!E2");
    expect(afterFirst).toBe("=(D2/A2)^(1/3)-1");

    helpers.select("Model!A2:D2");
    await smt.insertCagr();
    helpers.select("Model!A2:D2");
    await smt.insertCagr();
    expect(helpers.formula("Model!E2")).toBe(afterFirst);

    await smt.undoLastAction();
    await smt.undoLastAction();
    await smt.undoLastAction();
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});
