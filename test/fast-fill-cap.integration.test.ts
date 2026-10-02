// The fast fill cell cap: an extent sized off a large selection with nothing
// beside it (the "own" fallback in src/excel/formulas.ts) must be refused
// before the destination is ever read or written, the way autocolor and
// export already hold their own fills to SELECTION_CELL_CAP.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
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
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
});

describe("fast fill above the cell cap", () => {
  it("refuses a 6,000-cell extent before any grid load", async () => {
    helpers.seed("Model!A1", [[{ formula: "=1", value: 1 }]]);
    helpers.select("Model!A1:A6000");
    const before = helpers.syncCount();

    expect(await rejects(() => smt.fastFillAuto("down"))).toBe(
      "Fast fill supports up to 5,000 cells at once.",
    );
    // Only the three syncs that size the fill (the single-range check, the
    // active cell plus the selection, and the neighbour lines beside it): the
    // destination's own address, grid and write never ran.
    expect(helpers.syncCount() - before).toBe(3);
    expect(helpers.formula("Model!A6000")).toBe("");
  });

  it("still fills a 4,999-cell extent", async () => {
    helpers.seed("Model!A1", [[{ formula: "=1", value: 1 }]]);
    helpers.select("Model!A1:A4999");

    await smt.fastFillAuto("down");

    expect(helpers.formula("Model!A4999")).toBe("=1");
  });
});
