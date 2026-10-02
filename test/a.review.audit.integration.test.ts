// What the slice A review added against the strict fake host: the size cycles
// and the paintbrush answer a protected sheet with the staged sentence, and a
// sheet whose used range is too big to count (Range.cellCount answers -1) is
// skipped by the scans instead of read whole.

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

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  smt = await import("../src/excel");
});

describe("a protected sheet", () => {
  it("refuses the row height and column width cycles by name", async () => {
    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.applyRowHeightCycle())).toBe(
      "Row height: this sheet is protected, nothing was changed",
    );
    expect(await rejects(() => smt.applyColumnWidthCycle())).toBe(
      "Column width: this sheet is protected, nothing was changed",
    );
  });

  it("refuses a paint slot by name", async () => {
    helpers.seed("Data!A1", [[1]]);
    helpers.setActiveCell("Data!A1");
    const slot = await smt.captureSlot(1);

    helpers.seed("Model!A1", [[1]]);
    helpers.select("Model!A1");
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.applySlot(1, slot))).toBe(
      "paintbrush: this sheet is protected, nothing was changed",
    );
  });
});

describe("a sheet too big to count", () => {
  // A used range past 2^31-1 cells answers -1, which every ">" test read as
  // nothing at all: the scans then loaded the whole grid.
  it("is skipped by the model check", async () => {
    helpers.seed("Model!A1", [[{ formula: "=1+1", r1c1: "=1+1", value: 2 }]]);
    helpers.seed("Data!A1", [[1]]);
    helpers.seed("Data!XFD1048576", [[1]]);

    const { skipped } = await smt.runModelCheck();

    expect(skipped).toEqual(["Data"]);
  });
});
