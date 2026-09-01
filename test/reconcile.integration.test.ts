// The variance reconciliation adapter against the strict fake host: the
// numeric cells of one selected block feed the pure solver, and the matched
// cells become the selection so the modeller can inspect the answer in place.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  formatA1,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
});

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// 125, -40, 17, 300, -85, 9 in row-major order: the same set src/reconcile.test.ts
// proves reconciles to 49 via indices 0, 4 and 5 (A1, A3, B3). Row 4 mixes in
// the text and blank cells the solver must skip, and B2 is a formula so a
// computed value is read the same as a literal one.
function seedMixedBlock(): void {
  helpers.seed("Model!A1", [
    [125, -40],
    [17, { formula: "=SUM(A1:A1)", value: 300 }],
    [-85, 9],
    ["Total"],
  ]);
  helpers.select("Model!A1:B4");
}

describe("reconcileSelection", () => {
  it("selects the matched cells and reports sum, difference, addresses and values", async () => {
    seedMixedBlock();

    const result = await smt.reconcileSelection(49, 0.001);

    expect(result).toEqual({
      addresses: ["A1", "A3", "B3"],
      count: 3,
      difference: 0,
      sum: 49,
      values: [125, -85, 9],
    });
    const selected = workbook.selectionAreas.map((area) => formatA1(area.rect));
    expect(selected).toEqual(["A1", "A3", "B3"]);
    expect(
      workbook.selectionAreas.every(
        (area) => area.sheetId === helpers.sheet("Model").id,
      ),
    ).toBe(true);
  });

  it("takes a fixed number of round trips", async () => {
    seedMixedBlock();
    const before = helpers.syncCount();

    await smt.reconcileSelection(49, 0.001);

    expect(helpers.syncCount() - before).toBe(5);
  });

  // A clicked column header is a million cells: the cap answers before a
  // single value is read, the way the other selection tools refuse.
  it("refuses a selection over the cell cap before reading its values", async () => {
    helpers.select("Model!A1:A5001");

    expect(await rejects(() => smt.reconcileSelection(1, 0))).toBe(
      "Reconciliation supports up to 5,000 selected cells at once.",
    );
  });

  it("refuses a range with no numeric cell", async () => {
    helpers.seed("Model!A1", [["Total"], [""]]);
    helpers.select("Model!A1:A2");

    expect(await rejects(() => smt.reconcileSelection(1, 0))).toBe(
      "Select a range containing numeric cells.",
    );
  });

  it("refuses more than 34 numeric cells", async () => {
    helpers.seed(
      "Model!A1",
      Array.from({ length: 35 }, (_, i) => [i + 1]),
    );
    helpers.select("Model!A1:A35");

    expect(await rejects(() => smt.reconcileSelection(1, 0))).toBe(
      "Reconciliation supports up to 34 numeric cells; select a smaller range.",
    );
  });

  it("refuses a target no combination reaches within the tolerance", async () => {
    helpers.seed("Model!A1", [[1], [2], [4]]);
    helpers.select("Model!A1:A3");

    expect(await rejects(() => smt.reconcileSelection(20, 0))).toBe(
      "No combination reaches the target within the tolerance.",
    );
  });

  it("refuses a ctrl-clicked selection, named by stage", async () => {
    helpers.seed("Model!A1", [[1], [2]]);
    helpers.seed("Model!C1", [[3], [4]]);
    helpers.selectAreas(["Model!A1:A2", "Model!C1:C2"]);

    expect(await rejects(() => smt.reconcileSelection(5, 0))).toBe(
      "Reconciliation: select a single range",
    );
  });
});
