// Pass-2: Chart format over every Excel.ChartType value (node_modules/
// @types/office-js), "Invalid" excepted. styleChartLabels (src/excel/
// internal.ts) must set a label part only on a chart that has it - a wrong
// call ends the whole restyle on the web, not just a cosmetic batch.

import fc from "fast-check";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeChart,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

// Excel.ChartType's full value list, "Invalid" excepted: every chart type the
// active-chart guard (formatSelectedChart) can be asked to restyle, whether
// pls,fix drew it or the modeller inserted it through Excel's own ribbon.
const ALL_CHART_TYPES = [
  "ColumnClustered",
  "ColumnStacked",
  "ColumnStacked100",
  "3DColumnClustered",
  "3DColumnStacked",
  "3DColumnStacked100",
  "BarClustered",
  "BarStacked",
  "BarStacked100",
  "3DBarClustered",
  "3DBarStacked",
  "3DBarStacked100",
  "LineStacked",
  "LineStacked100",
  "LineMarkers",
  "LineMarkersStacked",
  "LineMarkersStacked100",
  "PieOfPie",
  "PieExploded",
  "3DPieExploded",
  "BarOfPie",
  "XYScatterSmooth",
  "XYScatterSmoothNoMarkers",
  "XYScatterLines",
  "XYScatterLinesNoMarkers",
  "AreaStacked",
  "AreaStacked100",
  "3DAreaStacked",
  "3DAreaStacked100",
  "DoughnutExploded",
  "RadarMarkers",
  "RadarFilled",
  "Surface",
  "SurfaceWireframe",
  "SurfaceTopView",
  "SurfaceTopViewWireframe",
  "Bubble",
  "Bubble3DEffect",
  "StockHLC",
  "StockOHLC",
  "StockVHLC",
  "StockVOHLC",
  "CylinderColClustered",
  "CylinderColStacked",
  "CylinderColStacked100",
  "CylinderBarClustered",
  "CylinderBarStacked",
  "CylinderBarStacked100",
  "CylinderCol",
  "ConeColClustered",
  "ConeColStacked",
  "ConeColStacked100",
  "ConeBarClustered",
  "ConeBarStacked",
  "ConeBarStacked100",
  "ConeCol",
  "PyramidColClustered",
  "PyramidColStacked",
  "PyramidColStacked100",
  "PyramidBarClustered",
  "PyramidBarStacked",
  "PyramidBarStacked100",
  "PyramidCol",
  "3DColumn",
  "Line",
  "3DLine",
  "3DPie",
  "Pie",
  "XYScatter",
  "3DArea",
  "Area",
  "Doughnut",
  "Radar",
  "Histogram",
  "Boxwhisker",
  "Pareto",
  "RegionMap",
  "Treemap",
  "Waterfall",
  "Sunburst",
  "Funnel",
] as const;

// Independent of src/excel/internal.ts's own (unexported) CHARTEX_TYPES: this
// is the doc-comment's list copied by hand, so a typo in one is not laundered
// by testing the other against itself.
const CHARTEX_TYPES = new Set([
  "Waterfall",
  "Treemap",
  "Sunburst",
  "Histogram",
  "Pareto",
  "Boxwhisker",
  "Funnel",
  "RegionMap",
]);

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"], web: true });
  helpers = host.helpers;
  smt = await import("../../src/excel");
});

// `prior` is what the fake seeded before the press: a part the chart type
// does not have is never assigned at all, so it must come back EXACTLY as it
// started (usually undefined - a fresh fake chart - but a hand-set true or
// false when a caller seeds one, which is the property test's whole point).
function expectLabelRule(
  chart: FakeChart,
  chartType: string,
  prior: {
    showPercentage?: boolean;
    showLegendKey?: boolean;
    showBubbleSize?: boolean;
  } = {},
): void {
  const isPieLike = /Pie|Doughnut/.test(chartType);
  const isBubble = chartType.startsWith("Bubble");
  const isChartex = CHARTEX_TYPES.has(chartType);

  expect(chart.dataLabels.showValue).toBe(true);
  expect(chart.dataLabels.showCategoryName).toBe(false);
  expect(chart.dataLabels.showSeriesName).toBe(false);

  if (isPieLike) expect(chart.dataLabels.showPercentage).toBe(false);
  else expect(chart.dataLabels.showPercentage).toBe(prior.showPercentage);

  if (isBubble) expect(chart.dataLabels.showBubbleSize).toBe(false);
  else expect(chart.dataLabels.showBubbleSize).toBe(prior.showBubbleSize);

  // showLegendKey is the one part every NON-chartex type has and every
  // chartex type refuses outright on the web (WEB_CHARTEX_REFUSED,
  // test/fakehost.ts): styleChartLabels sets it everywhere except the eight
  // chartex names, so a chartex chart must show no trace of the call at all.
  if (isChartex)
    expect(chart.dataLabels.showLegendKey).toBe(prior.showLegendKey);
  else expect(chart.dataLabels.showLegendKey).toBe(false);
}

describe("Chart format over every chart type the web can hold", () => {
  it("lands cleanly and sets only the label parts that type has, for every type", async () => {
    for (const chartType of ALL_CHART_TYPES) {
      for (const seriesCount of [1, 2, 6]) {
        const chart = helpers.addChart("Model", { chartType, seriesCount });
        helpers.setActiveChart(chart);

        await expect(smt.formatSelectedChart()).resolves.toBeUndefined();
        expectLabelRule(chart, chartType);
      }
    }
  });
});

describe("Chart format: the series count and any prior label state never change the rule", () => {
  const SEED = 20260927;

  it("holds under fast-check over every type, series count and starting label flags", async () => {
    const priorLabelsArb = fc.record({
      showPercentage: fc.boolean(),
      showLegendKey: fc.boolean(),
      showBubbleSize: fc.boolean(),
      showValue: fc.boolean(),
    });
    const caseArb = fc.record({
      chartType: fc.constantFrom(...ALL_CHART_TYPES),
      seriesCount: fc.integer({ min: 1, max: 8 }),
      prior: priorLabelsArb,
    });

    await fc.assert(
      fc.asyncProperty(caseArb, async ({ chartType, seriesCount, prior }) => {
        // Excel handed the chart back with whatever labels a previous,
        // possibly hand-made restyle left it in - Chart format must still
        // land on the one correct combination for this type, not merely
        // preserve or blindly toggle what was already there.
        const chart = helpers.addChart("Model", {
          chartType,
          seriesCount,
          dataLabels: { ...prior, font: {} },
        });
        helpers.setActiveChart(chart);

        await smt.formatSelectedChart();
        expectLabelRule(chart, chartType, prior);
      }),
      { seed: SEED, numRuns: 400 },
    );
  });
});
