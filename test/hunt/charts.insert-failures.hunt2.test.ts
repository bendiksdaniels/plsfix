// Pass-2 failure injection: every host sync refused, at every step, for the
// tornado, the football field, the waterfall and the CAGR arrow, desktop and
// web. Invariant: one plain sentence, no helper block or chart/shape left
// over the data, no Undo slot spent on a refusal, the next press works.

import { describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], ...options });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../../src/excel");
}

function seedGrowthLine(): void {
  helpers.seed("Model!A1", [[100, 110, 121, 133.1]]);
  helpers.select("Model!A1:D1");
}

function seedDrivers(): void {
  helpers.seed("Model!A1", [
    ["Driver", "Low", "High"],
    ["Volume", 90, 115],
    ["Price", 60, 140],
  ]);
  helpers.select("Model!A1:C3");
}

function seedMethods(): void {
  helpers.seed("Model!A1", [
    ["Method", "Low", "High"],
    ["DCF", 90, 130],
    ["Trading comps", 100, 120],
  ]);
  helpers.select("Model!A1:C3");
}

function seedBridge(): void {
  helpers.seed("Model!A1", [["EBITDA bridge"]]);
  helpers.seed("Model!A2", [
    ["Opening", 100],
    ["Price", 20],
    ["Cost", -30],
    ["Closing", 90],
  ]);
  helpers.select("Model!A2:B5");
}

async function expectNoBlockOrChart(): Promise<void> {
  // The helper block never landed, or was cleaned back off - a retry must
  // never meet its own leftovers instead of the sheet it actually selected.
  expect(helpers.value("Model!D1")).toBe("");
  expect(helpers.value("Model!E1")).toBe("");
  expect(helpers.value("Model!F1")).toBe("");
  expect(workbook.charts).toHaveLength(0);
  // A refusal that wrote nothing (or cleaned up what it wrote) never spends
  // a slot: nothing sits on the stack for this attempt to undo.
  await expect(smt.undoLastAction()).rejects.toThrow(
    "There is no pls,fix action to undo yet.",
  );
}

async function expectNoChart(): Promise<void> {
  // The waterfall writes no helper block - only a chart, if one landed at
  // all - so this is the narrower check insertWaterfall's own sweep needs.
  expect(workbook.charts).toHaveLength(0);
}

describe.each([{ web: false }, { web: true }])(
  "tornado: a refused sync at every step (web=$web)",
  ({ web }) => {
    it("for N = 0..last, refuses cleanly and leaves nothing behind, then the next press works", async () => {
      await boot({ web });
      seedDrivers();
      const before = helpers.syncCount();
      const clean = await smt.insertTornado();
      expect(clean).toContain("Tornado added");
      const total = helpers.syncCount() - before;

      for (let n = 0; n < total; n += 1) {
        await boot({ web });
        seedDrivers();
        helpers.failNextSync(undefined, n);

        await expect(smt.insertTornado()).rejects.toBeInstanceOf(Error);
        await expectNoBlockOrChart();

        const retry = await smt.insertTornado();
        expect(retry).toContain("Tornado added");
      }
    });
  },
);

describe("tornado: another pls,fix capture lands between the block write and the chart syncs", () => {
  // The pane has no busy latch over an Excel insert: a ribbon command or a
  // keyboard shortcut can capture another action while insertTornado sits
  // between its own block write and its chart syncs, pushing its own entry
  // on top of the stack. Simulated with a spy around the real
  // writeHelperBlock: once the tornado's own block lands, a second, real
  // capture-and-write runs on an unrelated range before tornado continues -
  // "a stub or spy around a real capture on another range", the reviewer's
  // own words for the simplest honest way to build this.
  let interleavePoint: number;

  async function installInterleave(): Promise<void> {
    const chartBlocks = await import("../../src/excel/chart-blocks");
    const undo = await import("../../src/excel/undo");
    const protection = await import("../../src/excel/protection");
    const real = chartBlocks.writeHelperBlock;
    let insertStartSyncs = 0;
    vi.spyOn(chartBlocks, "writeHelperBlock").mockImplementation(
      async (context, sheet, range, block) => {
        const target = await real(context, sheet, range, block);
        helpers.seed("Model!H1", [["old"], ["old"]]);
        helpers.select("Model!H1:H2");
        await Excel.run(async (otherContext) => {
          const other = otherContext.workbook.worksheets
            .getItem("Model")
            .getRange("H1:H2");
          await undo.captureUndo(otherContext, other);
          other.values = [["new"], ["new"]];
          await protection.syncWrite(otherContext, "Other action");
        });
        interleavePoint = helpers.syncCount() - insertStartSyncs;
        return target;
      },
    );
    insertStartSyncs = helpers.syncCount();
  }

  it("does not undo the other entry: its cells keep their values, it stays on top, the block is cleared", async () => {
    // Measure where the interleaved capture lands relative to insertTornado's
    // own start, on a clean run with nothing injected.
    await boot();
    seedDrivers();
    await installInterleave();
    const clean = await smt.insertTornado();
    expect(clean).toContain("Tornado added");

    // Same setup again, this time refusing the sync right after the
    // interleaved capture - insertTornado's own chart-add commit.
    await boot();
    seedDrivers();
    await installInterleave();
    helpers.failNextSync(undefined, interleavePoint);

    await expect(smt.insertTornado()).rejects.toBeInstanceOf(Error);

    expect(smt.undoTarget()).toBe("Model!H1:H2");
    expect(helpers.value("Model!H1")).toBe("new");
    expect(helpers.value("Model!H2")).toBe("new");
    expect(helpers.value("Model!D1")).toBe("");
    expect(workbook.charts).toHaveLength(0);
  });

  it("undoes the block's own entry as before when nothing else was captured in between", async () => {
    await boot();
    seedDrivers();
    await smt.insertTornado();
    const target = smt.undoTarget();

    const message = await smt.undoLastAction();
    expect(message).toBe(`Undone: ${target}. Nothing more to undo.`);
    expect(helpers.value("Model!D1")).toBe("");
  });
});

describe.each([{ web: false }, { web: true }])(
  "football field: a refused sync at every step (web=$web)",
  ({ web }) => {
    it("for N = 0..last, refuses cleanly and leaves nothing behind, then the next press works", async () => {
      await boot({ web });
      seedMethods();
      const before = helpers.syncCount();
      const clean = await smt.insertFootballField();
      expect(clean).toContain("Football field added");
      const total = helpers.syncCount() - before;

      for (let n = 0; n < total; n += 1) {
        await boot({ web });
        seedMethods();
        helpers.failNextSync(undefined, n);

        await expect(smt.insertFootballField()).rejects.toBeInstanceOf(Error);
        await expectNoBlockOrChart();

        const retry = await smt.insertFootballField();
        expect(retry).toContain("Football field added");
      }
    });
  },
);

describe.each([{ web: false }, { web: true }])(
  "waterfall: a refused sync at every step (web=$web)",
  ({ web }) => {
    it("for N = 0..last, refuses cleanly and leaves no chart behind, then the next press works", async () => {
      await boot({ web });
      seedBridge();
      const before = helpers.syncCount();
      const clean = await smt.insertWaterfall();
      expect(clean).toContain("Waterfall added");
      const total = helpers.syncCount() - before;

      for (let n = 0; n < total; n += 1) {
        await boot({ web });
        seedBridge();
        helpers.failNextSync(undefined, n);

        // A waterfall is a chartex chart, so on the web the surface batch
        // (chart.format.font/roundedCorners) refuses on its own, naturally,
        // at the same sync a run without any injection also reaches - and
        // that in-batch host error wins the same race an injected one does
        // (FakeContext.sync: an already-queued this.error is thrown ahead of
        // runtime.failSync, which is then cleared unfired). The injection is
        // lost, not merely tolerated, so this step can only be proven by its
        // outcome, not by an exception - already covered end to end by
        // web.waterfall.hunt.test.ts; here it only has to still mean a clean
        // insert with exactly one chart, never a silent partial one.
        let succeeded: string | undefined;
        try {
          succeeded = await smt.insertWaterfall();
        } catch (error) {
          expect(error).toBeInstanceOf(Error);
          await expectNoChart();
          const retry = await smt.insertWaterfall();
          expect(retry).toContain("Waterfall added");
          continue;
        }
        expect(succeeded).toContain("Waterfall added");
        expect(workbook.charts).toHaveLength(1);
      }
    });
  },
);

describe.each([{ web: false }, { web: true }])(
  "CAGR arrow: a refused sync at every step (web=$web)",
  ({ web }) => {
    it("for N = 0..last, refuses cleanly and leaves no shape behind, then the next press works", async () => {
      await boot({ web });
      seedGrowthLine();
      const before = helpers.syncCount();
      const clean = await smt.addCagrLabel();
      expect(clean).toBe("CAGR +10.0% over 3 periods");
      const total = helpers.syncCount() - before;

      for (let n = 0; n < total; n += 1) {
        await boot({ web });
        seedGrowthLine();
        helpers.failNextSync(undefined, n);

        await expect(smt.addCagrLabel()).rejects.toBeInstanceOf(Error);
        // The CAGR arrow is a shape, not a helper block: no cells to check,
        // and it never touches pls,fix Undo (shapes are worksheet objects,
        // outside the cell-format stack), only the shape itself.
        expect(workbook.shapes).toHaveLength(0);

        const retry = await smt.addCagrLabel();
        expect(retry).toBe("CAGR +10.0% over 3 periods");
      }
    });
  },
);
