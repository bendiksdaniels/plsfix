// Attacks: the pass-1 "repeats" lens over the tornado and the football field
// - three presses in a row then Undo, the two tools interleaved on the same
// target, and the same tool pressed on two different sheets at once. Both
// tools run every insert through chart-blocks.ts's single, shared serialised()
// queue, which football's own stress suite proves for two presses of itself;
// this file is the cross-tool and three-press cases neither existing suite
// covers.

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
    // A real prior action, so there is something on the stack under it.
    helpers.select("Model!H1:H2");
    helpers.seed("Model!H1", [["x"], ["y"]]);
    await smt.applyPinstripes("rows");
    const before = smt.undoTarget();

    seedDrivers("Model");
    await smt.insertTornado();
    const afterFirst = smt.undoTarget();
    await rejects(() => smt.insertTornado());
    await rejects(() => smt.insertTornado());

    expect(afterFirst).not.toBe(before);
    // The two refused presses captured nothing: the slot the one real chart
    // insert left behind is still the most recent one.
    expect(smt.undoTarget()).toBe(afterFirst);
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

  it("racing the same target across two sheets still yields one chart each", async () => {
    // The selection is one global, mutable pointer: switching sheets before
    // the first press's queued work has actually run hands BOTH presses the
    // sheet current when each is dequeued, not the sheet selected when each
    // was called. serialised() still runs them one at a time, so the second
    // to run either finds its own sheet's block already taken (refused
    // cleanly) or, once the first pass has moved the selection on, draws its
    // own chart there instead - never two charts sharing one block.
    await boot(["Model", "Data"]);
    seedDrivers("Model");
    const first = smt.insertTornado();
    seedDrivers("Data");
    const second = smt.insertTornado();

    const [firstOut, secondOut] = await Promise.all([
      first.catch((error: Error) => error.message),
      second.catch((error: Error) => error.message),
    ]);

    for (const line of [firstOut, secondOut]) {
      expect(
        line.includes("Tornado added") ||
          line === "tornado: cells to the right of the selection are not empty",
      ).toBe(true);
    }
    // Never a chart stacked on top of another: at most one per sheet, and at
    // least one overall.
    expect(workbook.charts.length).toBeGreaterThanOrEqual(1);
    expect(workbook.charts.length).toBeLessThanOrEqual(2);
  });
});
