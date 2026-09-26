// Attacks: Excel for the web's refusal of a chart's cosmetic surface
// (chart.format.font, roundedCorners) on the tornado and football-field
// paths - both UnsupportedOperation (chartex charts) and InvalidOperation
// ("...not permitted for the current object.", the code public traces show).

import { beforeEach, describe, expect, it, vi } from "vitest";
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

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// Two drivers at the very corner: the same shape tornado.integration.test.ts
// uses for its plain-host base case ("Tornado added: 2 drivers, base 101.3").
function seedDrivers(): void {
  helpers.seed("Model!A1", [
    ["Driver", "Low", "High"],
    ["Volume", 90, 115],
    ["Price", 60, 140],
  ]);
  helpers.select("Model!A1:C3");
}

// Two valuation methods with a header row, the same shape football's own
// suite seeds.
function seedMethods(): void {
  helpers.seed("Model!A1", [
    ["Method", "Low", "High"],
    ["DCF", 90, 130],
    ["Trading comps", 100, 120],
  ]);
  helpers.select("Model!A1:C3");
}

beforeEach(async () => {
  await boot();
});

describe("the tornado when Excel for the web refuses the chart surface", () => {
  it("lands the chart exactly where it would without the refusal", async () => {
    // Same selection, same host otherwise: only the surface call differs.
    await boot({ chartSurfaceUnsupported: false });
    seedDrivers();
    await smt.insertTornado();
    const clean = workbook.charts[0];
    const cleanCorner = { left: clean?.left, top: clean?.top };

    await boot({ chartSurfaceUnsupported: true });
    seedDrivers();
    await smt.insertTornado();
    const refused = workbook.charts[0];

    expect({ left: refused?.left, top: refused?.top }).toEqual(cleanCorner);
    // The measured corner right of the helper block, never Excel's own drop
    // point at the origin, which sits over the model.
    expect(refused).toMatchObject({ left: 7 * 64, top: 0 });
  });

  it("answers the normal sentence with no trace of the tolerated refusal", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedDrivers();

    const message = await smt.insertTornado();

    expect(message).toBe("Tornado added: 2 drivers, base 101.3");
  });

  it("keeps every branding step that is not the refused surface", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedDrivers();

    await smt.insertTornado();

    const chart = workbook.charts[0];
    expect(chart?.title).toBe("Sensitivity");
    expect(chart?.dataLabels).toMatchObject({
      showValue: true,
      position: "OutsideEnd",
    });
    expect(chart?.axes.category.reversePlotOrder).toBe(true);
    expect(chart?.series[0]?.overlap).toBe(100);
    // The surface itself: refused, so never written.
    expect(chart?.font).toEqual({});
    expect(chart?.roundedCorners).toBeUndefined();
  });

  it("presses three times: one placed chart, two clean refusals, never a stack over the data", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedDrivers();

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
    expect(workbook.charts[0]).toMatchObject({ left: 7 * 64, top: 0 });
  });
});

describe("the football field when Excel for the web refuses the chart surface", () => {
  it("lands the chart exactly where it would without the refusal", async () => {
    await boot({ chartSurfaceUnsupported: false });
    seedMethods();
    await smt.insertFootballField();
    const clean = workbook.charts[0];
    const cleanCorner = { left: clean?.left, top: clean?.top };

    await boot({ chartSurfaceUnsupported: true });
    seedMethods();
    await smt.insertFootballField();
    const refused = workbook.charts[0];

    expect({ left: refused?.left, top: refused?.top }).toEqual(cleanCorner);
    expect(refused).toMatchObject({ left: 7 * 64, top: 0 });
  });

  it("answers the normal sentence with no trace of the tolerated refusal", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedMethods();

    const message = await smt.insertFootballField();

    expect(message).toBe("Football field added: 2 ranges");
  });

  it("keeps every branding step that is not the refused surface", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedMethods();

    await smt.insertFootballField();

    const chart = workbook.charts[0];
    expect(chart?.legend.visible).toBe(false);
    expect(chart?.axes.category.reversePlotOrder).toBe(true);
    expect(chart?.series[0]).toMatchObject({ fillCleared: true });
    expect(chart?.font).toEqual({});
    expect(chart?.roundedCorners).toBeUndefined();
  });

  it("presses three times: one placed chart, two clean refusals, never a stack over the data", async () => {
    await boot({ chartSurfaceUnsupported: true });
    seedMethods();

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
    expect(workbook.charts[0]).toMatchObject({ left: 7 * 64, top: 0 });
  });
});

// The ledger's own capture showed the web naming a different code than the
// chartex refusal the fake defaults to: InvalidOperation and "This operation
// is not permitted for the current object.", via the opt-in override.
describe("a surface refusal coded InvalidOperation instead of UnsupportedOperation", () => {
  function bootInvalidOperation(): Promise<void> {
    return boot({
      chartSurfaceUnsupported: true,
      chartSurfaceRefusal: {
        code: Excel.ErrorCodes.invalidOperation,
        message: "This operation is not permitted for the current object.",
      },
    });
  }

  it("still places the tornado beside the data with a clean toast", async () => {
    await bootInvalidOperation();
    seedDrivers();

    const message = await smt.insertTornado();

    expect(message).toBe("Tornado added: 2 drivers, base 101.3");
    expect(workbook.charts[0]).toMatchObject({ left: 7 * 64, top: 0 });
  });

  it("still brands a bar chart through the chart-format restyle", async () => {
    await bootInvalidOperation();
    seedDrivers();
    await smt.insertTornado();
    const chart = workbook.charts[0];
    if (!chart) throw new Error("no chart");
    helpers.setActiveChart(chart);

    await smt.formatSelectedChart();

    expect(chart.legend.position).toBe("Bottom");
    expect(chart.font).toEqual({});
    expect(chart.roundedCorners).toBeUndefined();
  });
});
