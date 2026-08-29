// Every tool on the selection a modeller actually makes rather than the one the
// tool was written for: a bridge laid out sideways, a growth line with blanks at
// its ends, a rounding block, a fast fill with nothing beside it, and a cross
// tab with gaps and headers that are not text.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
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

beforeEach(async () => {
  await boot();
});

describe("a bridge in two adjacent rows", () => {
  // Labels along the top, values under them, the heading to their left.
  function seedRowBridge(): void {
    helpers.seed("Model!A2", [["EBITDA bridge"]]);
    helpers.seed("Model!B2", [
      ["Opening", "Price", "Cost", "Closing"],
      [100, 20, -30, 90],
    ]);
    helpers.select("Model!B2:E3");
  }

  it("reads it the way it reads two columns", async () => {
    seedRowBridge();
    expect(await smt.insertWaterfall()).toBe(
      "Waterfall added: 4 points, ties at 90",
    );
    expect(workbook.charts[0]).toMatchObject({
      chartType: "Waterfall",
      sourceAddress: "Model!B2:E3",
      seriesBy: "Rows",
      title: "EBITDA bridge",
    });
  });

  it("brands the totals and the fall by position", async () => {
    seedRowBridge();
    await smt.insertWaterfall();
    const colors = workbook.charts[0]?.series[0]?.pointColors;
    expect(Object.keys(colors ?? {})).toEqual(["0", "1", "2", "3"]);
  });

  it("names the row that has to hold numbers", async () => {
    helpers.seed("Model!A1", [
      ["Opening", "Price", "Closing"],
      [100, "n/a", 120],
    ]);
    helpers.select("Model!A1:C2");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "The second row must hold numbers only.",
    );
  });

  it("still refuses a block that is neither", async () => {
    helpers.select("Model!A1:C4");
    expect(await rejects(() => smt.insertWaterfall())).toBe(
      "Select labels and values in two adjacent columns, or in two adjacent rows, with three or more points.",
    );
  });
});

describe("a growth line with blanks at its ends", () => {
  it("writes the CAGR just past the last number in a row", async () => {
    helpers.seed("Model!A2", [["", 100, 110, 121, ""]]);
    helpers.select("Model!A2:E2");

    await smt.insertCagr();

    expect(helpers.formula("Model!E2")).toBe("=(D2/B2)^(1/2)-1");
    expect(helpers.numberFormat("Model!E2")).toBe("0.0%;[Red](0.0%);-");
  });

  it("writes it just below the last number in a column", async () => {
    helpers.seed("Model!B1", [[""], [100], [121], [""]]);
    helpers.select("Model!B1:B4");

    await smt.insertCagr();
    expect(helpers.formula("Model!B4")).toBe("=(B3/B2)^(1/1)-1");
  });

  it("labels a line whose ends are blank", async () => {
    helpers.seed("Model!A1", [["", 100, 110, 121, 133.1, ""]]);
    helpers.select("Model!A1:F1");

    expect(await smt.addCagrLabel()).toBe("CAGR +10.0% over 3 periods");
  });

  it("still needs two numbers", async () => {
    helpers.seed("Model!A1", [["", 100, ""]]);
    helpers.select("Model!A1:C1");
    expect(await rejects(() => smt.insertCagr())).toBe(
      "Select one row or column with at least two periods.",
    );
  });
});

describe("consistent rounding on a line", () => {
  it("takes a row and a column, and names a block as a block", async () => {
    helpers.seed("Model!B2", [[33.333, 33.333, 33.334]]);
    helpers.select("Model!B2:D2");
    expect(await smt.insertConsistentRounding()).toBe(
      "Consistent rounding: 3 cells at 0 decimals",
    );

    helpers.seed("Model!B6", [[33.333], [33.333], [33.334]]);
    helpers.select("Model!B6:B8");
    expect(await smt.insertConsistentRounding()).toBe(
      "Consistent rounding: 3 cells at 0 decimals",
    );

    helpers.seed("Model!F1", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!F1:G2");
    expect(await rejects(() => smt.insertConsistentRounding())).toBe(
      "Consistent rounding: select a single row or a single column, not a block.",
    );
  });
});

describe("a fast fill with nothing beside it", () => {
  it("fills the selection's own width", async () => {
    helpers.seed("Model!C5", [[{ formula: "=C4*2", value: 1 }]]);
    helpers.select("Model!C5:F5");

    await smt.fastFillAuto("right");

    expect(helpers.formula("Model!F5")).toBe("=C4*2");
    expect(helpers.formula("Model!G5")).toBe("");
  });

  it("fills the selection's own height", async () => {
    helpers.seed("Model!C5", [[{ formula: "=C4*2", value: 1 }]]);
    helpers.select("Model!C5:C8");

    await smt.fastFillAuto("down");

    expect(helpers.formula("Model!C8")).toBe("=C4*2");
    expect(helpers.formula("Model!C9")).toBe("");
  });

  it("still refuses a lone cell with nothing beside it", async () => {
    helpers.seed("Model!C5", [[{ formula: "=C4*2", value: 1 }]]);
    helpers.select("Model!C5");

    expect(await rejects(() => smt.fastFillAuto("down"))).toBe(
      "No neighbor data to size the fill.",
    );
  });

  it("lets the neighbour data win when there is any", async () => {
    helpers.seed("Model!B1", [["Q1", "Q2", "Q3", "Q4"]]);
    helpers.seed("Model!B2", [[{ formula: "=A2*2", value: 8 }]]);
    helpers.select("Model!B2:C2");

    await smt.fastFillAuto("right");

    expect(helpers.formula("Model!E2")).toBe("=A2*2");
  });
});

describe("unpivot over an untidy cross tab", () => {
  it("skips the gaps and keeps headers that are not text", async () => {
    helpers.seed("Model!A1", [
      ["", 2024, 2025, true],
      ["Revenue", 10, "", 1],
      ["", 99, 99, 99],
      ["Costs", "", 20, 2],
    ]);
    helpers.select("Model!A1:D4");

    expect(await smt.unpivotSelection()).toBe("Unpivot: 4 rows on Unpivot");

    const rows = ["A2", "B2", "C2"].map((at) => helpers.value(`Unpivot!${at}`));
    expect(rows).toEqual(["Revenue", 2024, 10]);
    expect(
      ["A3", "B3", "C3"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Revenue", true, 1]);
    expect(
      ["A4", "B4", "C4"].map((at) => helpers.value(`Unpivot!${at}`)),
    ).toEqual(["Costs", 2025, 20]);
  });

  it("says so when the gaps are all there is", async () => {
    helpers.seed("Model!A1", [
      ["", "North"],
      ["Q1", ""],
    ]);
    helpers.select("Model!A1:B2");

    expect(await rejects(() => smt.unpivotSelection())).toBe(
      "unpivot: the selection holds no values",
    );
  });
});
