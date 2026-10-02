// The names scrubber over sheet-scoped defined names (Worksheet.names,
// ExcelApi 1.4): workbook.names alone missed a name someone scoped to one
// sheet, so a broken one there never showed up to list or delete. Driven end
// to end against the strict fake host; the workbook-scoped path itself is
// covered in guards.integration.test.ts and stays untouched here.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  hostError,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
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
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
});

describe("sheet-scoped broken names", () => {
  beforeEach(() => {
    helpers
      .sheet("Model")
      .names.push({ name: "TaxRate", formula: "=Model!#REF!", visible: true });
    helpers
      .sheet("Data")
      .names.push({ name: "Region", formula: "=Data!$A$1", visible: true });
  });

  it("lists a broken one as Sheet!Name, leaves a healthy one out", async () => {
    expect(await smt.listBrokenNames()).toEqual(["Model!TaxRate"]);
  });

  it("deletes a broken one and counts it, keeping the healthy one", async () => {
    expect(await smt.deleteBrokenNames()).toBe(1);

    expect(helpers.sheet("Model").names.map((entry) => entry.name)).toEqual([]);
    expect(helpers.sheet("Data").names.map((entry) => entry.name)).toEqual([
      "Region",
    ]);
  });

  it("still lists and deletes a workbook-scoped broken name unprefixed", async () => {
    helpers.addName("Costs", "=Model!#REF!");

    expect(await smt.listBrokenNames()).toEqual(["Costs", "Model!TaxRate"]);
    expect(await smt.deleteBrokenNames()).toBe(2);
    expect(workbook.names).toEqual([]);
    expect(helpers.sheet("Model").names).toEqual([]);
  });
});

describe("deleting broken names on a protected workbook", () => {
  it("answers a pane sentence, not Excel's own string", async () => {
    helpers.addName("Costs", "=Model!#REF!");
    // deleteBrokenNames runs two loading syncs (the sheet list plus the
    // workbook's names, then every sheet's own names) before the delete
    // write; afterSyncs lets those two succeed and fails only the write.
    helpers.failNextSync(
      hostError("AccessDenied", "The workbook is protected."),
      2,
    );

    expect(await rejects(() => smt.deleteBrokenNames())).toBe(
      "The workbook's structure is protected, so nothing was deleted.",
    );
  });
});
