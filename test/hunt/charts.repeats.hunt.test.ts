// Attacks: the pass-1 "repeats" lens over the tornado and the football field -
// three presses in a row then Undo, the two tools interleaved on the same
// target, and the same tool raced across two sheets at once, all through
// chart-blocks.ts's single, shared serialised() queue.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(sheets: string[] = ["Model"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  helpers = host.helpers;
  workbook = host.workbook;
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

function seedDrivers(sheet: string): void {
  helpers.seed(`${sheet}!A1`, [
    ["Driver", "Low", "High"],
    ["Volume", 90, 115],
    ["Price", 60, 140],
  ]);
  helpers.select(`${sheet}!A1:C3`);
}

function seedMethods(sheet: string): void {
  helpers.seed(`${sheet}!A1`, [
    ["Method", "Low", "High"],
    ["DCF", 90, 130],
    ["Trading comps", 100, 120],
  ]);
  helpers.select(`${sheet}!A1:C3`);
}

beforeEach(async () => {
  await boot();
});

describe("the tornado pressed three times in a row", () => {
  it("draws one chart, refuses the other two cleanly, then Undo removes it", async () => {
    seedDrivers("Model");

    const first = await smt.insertTornado();
    const second = await rejects(() => smt.insertTornado());
    const third = await rejects(() => smt.insertTornado());

    expect(first).toContain("Tornado added: 2 drivers");
    expect(second).toBe(
      "tornado: cells to the right of the selection are not empty",
    );
    expect(third).toBe(
      "tornado: cells to the right of the selection are not empty",
    );
    expect(workbook.charts).toHaveLength(1);
    expect(helpers.value("Model!D1")).toBe("Driver");

    await smt.undoLastAction();
    expect(helpers.value("Model!D1")).toBe("");
  });

  it("spends exactly one Undo slot: the two refusals never capture one", async () => {
    // Fill the 5-deep stack first: with slack left in it, a refusal that
    // wrongly captured a slot and then discarded it again as a stale pending
    // entry (undoLastAction's own discardUndo runs before it reads the top)
    // would leave no trace, since every capture here targets the same
    // address. Full, the same bug evicts a filler from the bottom that never
    // comes back - visible as one fewer entry left after undoing the tornado.
    for (const range of ["H1:H2", "H3:H4", "H5:H6", "H7:H8", "H9:H10"]) {
      helpers.select(`Model!${range}`);
      helpers.seed(`Model!${range.split(":")[0]}`, [["x"], ["y"]]);
      await smt.applyPinstripes("rows");
    }

    seedDrivers("Model");
    await smt.insertTornado();
    const afterFirst = smt.undoTarget();
    await rejects(() => smt.insertTornado());
    await rejects(() => smt.insertTornado());

    expect(smt.undoTarget()).toBe(afterFirst);
    // The full-stack proof: undoing the tornado reveals exactly four more
    // entries. A refusal that had wrongly spent and evicted one from the
    // bottom would leave three, even though its own entry is discarded
    // before this message is built.
    const message = await smt.undoLastAction();
    expect(message).toBe(`Undone: ${afterFirst}. 4 more to undo.`);
  });
});

describe("the football field pressed three times in a row", () => {
  it("draws one chart and refuses the other two cleanly", async () => {
    seedMethods("Model");

    const first = await smt.insertFootballField();
    const second = await rejects(() => smt.insertFootballField());
    const third = await rejects(() => smt.insertFootballField());

    expect(first).toContain("Football field added: 2 ranges");
    expect(second).toBe(
      "Football field: cells to the right of the selection are not empty",
    );
    expect(third).toBe(
      "Football field: cells to the right of the selection are not empty",
    );
    expect(workbook.charts).toHaveLength(1);
  });
});

describe("the tornado and the football field interleaved on the same target", () => {
  it("lets exactly one land and refuses the other, never two charts on one block", async () => {
    seedDrivers("Model");

    const [tornado, football] = await Promise.all([
      smt.insertTornado().catch((error: Error) => error.message),
      smt.insertFootballField().catch((error: Error) => error.message),
    ]);

    // Both run through chart-blocks.ts's one shared serialised() queue: the
    // second to reach writeHelperBlock meets the first's block already
    // sitting where its own would go, and is refused by name, not silently
    // merged or left to draw a second chart on top of the first.
    const outcomes = [tornado, football];
    const succeeded = outcomes.filter((line) => line.includes("added"));
    const refused = outcomes.filter((line) =>
      line.includes("cells to the right of the selection are not empty"),
    );
    expect(succeeded).toHaveLength(1);
    expect(refused).toHaveLength(1);
    expect(workbook.charts).toHaveLength(1);
  });
});

describe("the same tool pressed on two different sheets", () => {
  it("draws a tornado on each sheet with no state left over from the first", async () => {
    await boot(["Model", "Data"]);
    seedDrivers("Model");
    const onModel = await smt.insertTornado();

    seedDrivers("Data");
    const onData = await smt.insertTornado();

    expect(onModel).toContain("Tornado added: 2 drivers");
    expect(onData).toContain("Tornado added: 2 drivers");
    expect(workbook.charts).toHaveLength(2);
    expect(helpers.value("Model!D1")).toBe("Driver");
    expect(helpers.value("Data!D1")).toBe("Driver");
  });

  it("racing the same target across two sheets: the first queued wins, the second is refused", async () => {
    // The selection is one global, mutable pointer: switching sheets before
    // the first press's queued work has actually run hands BOTH presses the
    // sheet current when each is dequeued, not the sheet selected when each
    // was called - so both race over "Data". serialised() still runs them
    // one at a time: the first queued writes the block and draws its chart,
    // the second meets that block already written and is refused by name.
    // With serialised() removed this stays green while two charts stack on
    // Data!D1:F3, so the outcome asserted below has to be the exact one
    // serialisation guarantees, not just "one or two, somewhere".
    await boot(["Model", "Data"]);
    seedDrivers("Model");
    const first = smt.insertTornado();
    seedDrivers("Data");
    const second = smt.insertTornado();

    const [firstOut, secondOut] = await Promise.all([
      first.catch((error: Error) => error.message),
      second.catch((error: Error) => error.message),
    ]);

    expect(firstOut).toContain("Tornado added");
    expect(secondOut).toBe(
      "tornado: cells to the right of the selection are not empty",
    );
    // Exactly one chart, on the sheet both presses raced over.
    expect(workbook.charts.map((chart) => chart.sheetName)).toEqual(["Data"]);
  });
});
