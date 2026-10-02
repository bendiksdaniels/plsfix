// The football field against the strict fake host: the helper block beside the
// selection, the stacked bar built off it, the swap a backwards row gets, and
// every selection a modeller can hand it - multi-area, merged, protected, over
// the cap and empty.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import { seriesPalette } from "../src/chart-colors";
import { DEFAULT_SETTINGS } from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
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

// Three valuation methods with a header row at A1:C4.
function seedMethods(): void {
  helpers.seed("Model!A1", [
    ["Method", "Low", "High"],
    ["DCF", 90, 130],
    ["Trading comps", 100, 120],
    ["Precedents", 110, 150],
  ]);
  helpers.select("Model!A1:C4");
}

beforeEach(async () => {
  await boot();
});

describe("insertFootballField", () => {
  it("writes a label, low and range block beside the selection", async () => {
    seedMethods();

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 3 ranges",
    );

    expect(
      ["D1", "E1", "F1"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Method", "Low", "Range"]);
    expect(
      ["D2", "E2", "F2"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["DCF", 90, 40]);
    expect(
      ["D4", "E4", "F4"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Precedents", 110, 40]);
  });

  it("builds a stacked bar with the floor hidden and the band branded", async () => {
    seedMethods();
    await smt.insertFootballField();

    const chart = workbook.charts[0];
    expect(chart).toMatchObject({
      chartType: "BarStacked",
      seriesBy: "Columns",
      sourceAddress: "Model!D1:F4",
      title: "Valuation range",
    });
    expect(chart?.series[0]).toMatchObject({
      fillCleared: true,
      lineStyle: "None",
    });
    expect(chart?.series[1]?.fillColor).toBe(
      seriesPalette(DEFAULT_SETTINGS)[0],
    );
    expect(chart?.legend.visible).toBe(false);
    expect(chart?.axes.category.majorGridlines).toBe(false);
    expect(chart?.axes.value.majorGridlines).toBe(false);
  });

  it("puts the first row on top and labels the axis like the low column", async () => {
    seedMethods();
    helpers.setNumberFormat("Model!B2:C4", "#,##0");
    await smt.insertFootballField();

    const chart = workbook.charts[0];
    expect(chart?.axes.category.reversePlotOrder).toBe(true);
    expect(chart?.axes.value.numberFormat).toBe("#,##0");
    // The helper block's own numbers wear the same format.
    expect(helpers.numberFormat("Model!E2")).toBe("#,##0");
    expect(helpers.numberFormat("Model!F2")).toBe("#,##0");
    expect(helpers.numberFormat("Model!D2")).toBe("General");
  });

  it("swaps a row entered high first and says how many", async () => {
    helpers.seed("Model!A2", [
      ["DCF", 130, 90],
      ["Comps", 120, 100],
      ["Precedents", 110, 150],
    ]);
    helpers.select("Model!A2:C4");

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 3 ranges; 2 rows had low above high, swapped",
    );
    // The block starts on the selection's own first row, so row 2 is its
    // header and row 3 the first method.
    expect(helpers.value("Model!E3")).toBe(90);
    expect(helpers.value("Model!F3")).toBe(40);
  });

  it("says one row in the singular", async () => {
    helpers.seed("Model!A2", [
      ["DCF", 130, 90],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!A2:C3");

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 2 ranges; 1 row had low above high, swapped",
    );
  });

  it("ignores a fourth column and says so", async () => {
    helpers.seed("Model!A1", [
      ["Method", "Low", "High", "Point"],
      ["DCF", 90, 130, 110],
      ["Comps", 100, 120, 110],
    ]);
    helpers.select("Model!A1:D3");

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 2 ranges; only label, low and high are used",
    );
    // The helper block starts past the whole selection, four columns wide.
    expect(helpers.value("Model!E1")).toBe("Method");
  });

  it("puts the helper block back on Undo", async () => {
    seedMethods();
    await smt.insertFootballField();
    await smt.undoLastAction();

    expect(helpers.value("Model!D1")).toBe("");
    expect(helpers.value("Model!E2")).toBe("");
  });
});

describe("what the football field refuses", () => {
  it("refuses fewer than three columns", async () => {
    helpers.seed("Model!A1", [
      ["DCF", 90],
      ["Comps", 100],
    ]);
    helpers.select("Model!A1:B2");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: select at least three columns - label, low and high.",
    );
  });

  it("refuses fewer than two rows", async () => {
    helpers.seed("Model!A1", [
      ["Method", "Low", "High"],
      ["DCF", 90, 130],
    ]);
    helpers.select("Model!A1:C2");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: need at least two rows",
    );
  });

  it("refuses more rows than the cap", async () => {
    const rows = Array.from({ length: 21 }, (_unused, index) => [
      `Method ${String(index)}`,
      90,
      130,
    ]);
    helpers.seed("Model!A1", rows);
    helpers.select("Model!A1:C21");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field supports up to 20 rows.",
    );
  });

  it("refuses a low or a high that is not a number", async () => {
    helpers.seed("Model!A1", [
      ["DCF", 90, "n/a"],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!A1:C2");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: the low and high columns must hold numbers",
    );
  });

  it("refuses an empty selection the same way", async () => {
    helpers.select("Model!C3:E6");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: the low and high columns must hold numbers",
    );
  });

  it("refuses a ctrl-clicked selection", async () => {
    seedMethods();
    helpers.selectAreas(["Model!A1:C4", "Model!H1:J2"]);

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: select a single range",
    );
  });

  it("refuses a selection over the cell cap, before any grid is read", async () => {
    helpers.select("Model!A:C");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field supports up to 5,000 selected cells at once.",
    );
  });

  it("refuses a block with no room to its right", async () => {
    helpers.seed("Model!XEZ1", [
      ["DCF", 90, 130],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!XEZ1:XFB2");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: no room to the right of the selection",
    );
  });

  it("refuses to overwrite what stands beside the selection", async () => {
    seedMethods();
    helpers.seed("Model!E2", [["busy"]]);

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: cells to the right of the selection are not empty",
    );
  });

  it("names the protected sheet instead of Excel's own string", async () => {
    seedMethods();
    helpers.protectSheet("Model");

    expect(await rejects(() => smt.insertFootballField())).toBe(
      "Football field: this sheet is protected, nothing was changed",
    );
    expect(workbook.charts).toHaveLength(0);
  });

  it("spends no Undo slot on the protected-sheet refusal", async () => {
    // One real action first, so there is something on the stack to lose.
    helpers.select("Model!A1:C2");
    await smt.applyPinstripes("rows");
    const slot = smt.undoTarget();

    seedMethods();
    helpers.protectSheet("Model");
    await rejects(() => smt.insertFootballField());

    expect(smt.undoTarget()).toBe(slot);
  });
});

describe("the football field on a merged selection and an older host", () => {
  it("reads a merged heading band as the header row", async () => {
    helpers.seed("Model!A1", [["Valuation summary"]]);
    helpers.merge("Model!A1:C1");
    helpers.seed("Model!A2", [
      ["DCF", 90, 130],
      ["Comps", 100, 120],
    ]);
    helpers.select("Model!A1:C3");

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 2 ranges",
    );
    expect(helpers.value("Model!D2")).toBe("DCF");
  });

  it("says the axes are plain when the host has neither 1.7 nor 1.8", async () => {
    helpers.setSupported(
      (_set, version) => version !== "1.7" && version !== "1.8",
    );
    seedMethods();

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 3 ranges; plain axes on this build",
    );
    expect(workbook.charts[0]?.axes.category.reversePlotOrder).toBeUndefined();
    expect(workbook.charts[0]?.axes.value.numberFormat).toBeUndefined();
  });
});
