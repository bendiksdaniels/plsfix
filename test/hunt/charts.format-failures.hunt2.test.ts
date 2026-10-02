// Pass-2 failure injection: every host sync refused, at every step, for
// Chart format over column/bar/line/pie/doughnut/waterfall/stacked, desktop
// and web. No helper block, no Undo slot, nothing new created - so the only
// invariant is a clean error and a fully working next press.

import { describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeChart,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], ...options });
  helpers = host.helpers;
  smt = await import("../../src/excel");
}

const NAMED_TYPES: { name: string; chartType: string; seriesCount: number }[] =
  [
    { name: "column", chartType: "ColumnClustered", seriesCount: 2 },
    { name: "bar", chartType: "BarClustered", seriesCount: 2 },
    { name: "line", chartType: "Line", seriesCount: 2 },
    { name: "pie", chartType: "Pie", seriesCount: 1 },
    { name: "doughnut", chartType: "Doughnut", seriesCount: 1 },
    { name: "waterfall", chartType: "Waterfall", seriesCount: 1 },
    { name: "stacked", chartType: "ColumnStacked", seriesCount: 2 },
  ];

describe.each([{ web: false }, { web: true }])(
  "Chart format: a refused sync at every step (web=$web)",
  ({ web }) => {
    for (const { name, chartType, seriesCount } of NAMED_TYPES) {
      it(`${name}: for N = 0..last, refuses cleanly and stays usable, then the next press restyles fully`, async () => {
        await boot({ web });
        const clean = helpers.addChart("Model", { chartType, seriesCount });
        helpers.setActiveChart(clean);
        const before = helpers.syncCount();
        await smt.formatSelectedChart();
        const total = helpers.syncCount() - before;

        for (let n = 0; n < total; n += 1) {
          await boot({ web });
          const chart = helpers.addChart("Model", { chartType, seriesCount });
          helpers.setActiveChart(chart);
          helpers.failNextSync(undefined, n);

          // A chartex chart (the waterfall) refuses its own surface on the
          // web naturally, at the same sync a run with no injection also
          // reaches - the in-batch host error wins the race an injected one
          // does (FakeContext.sync: an already-queued this.error is thrown
          // ahead of runtime.failSync, which is then cleared unfired), so
          // the injection is lost rather than merely tolerated. That step
          // can only be proven by its outcome: a chart still restyled
          // exactly as a clean run would leave it, never a silent partial
          // one, already covered end to end for that one type by
          // web.charts.hunt.test.ts.
          let threw = false;
          try {
            await smt.formatSelectedChart();
          } catch (error) {
            threw = true;
            expect(error).toBeInstanceOf(Error);
          }
          if (!threw) {
            expectFullyStyled(chart, chartType);
          }

          // Whatever this attempt did to the chart, Chart format never
          // blocks a later press the way an insert's "not empty" can: the
          // same chart, restyled again, always lands the full result.
          await smt.formatSelectedChart();
          expectFullyStyled(chart, chartType);
        }
      });
    }
  },
);

function expectFullyStyled(chart: FakeChart, chartType: string): void {
  expect(chart.dataLabels.showValue).toBe(true);
  expect(chart.dataLabels.showCategoryName).toBe(false);
  expect(chart.dataLabels.showSeriesName).toBe(false);
  if (/Pie|Doughnut/.test(chartType)) {
    expect(chart.dataLabels.showPercentage).toBe(false);
  }
  expect(chart.legend.font.name).toBeTruthy();
}
