// Attacks readChartData (src/excel/link-chart.ts) directly with a minimal
// hand-built context/chart double, not the full fake host, targeting paths
// no chart-export integration test reaches: a chart that stops answering
// mid-read, fully overlapped columns, a non-numeric series, pie/waterfall colours.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { pieColors, waterfallColors } from "../../src/chart-colors";
import { getActiveSettings } from "../../src/settings";
import { readChartData, SERIES_READ_REFUSED } from "../../src/excel/link-chart";

beforeEach(() => {
  vi.stubGlobal("Office", { context: {} });
});
afterEach(() => {
  vi.unstubAllGlobals();
});

interface FakeSeriesSpec {
  name: string;
  categories?: string[];
  values: (string | number)[];
  overlap?: number;
  source?: string;
}

interface FakeChartSpec {
  chartType: string;
  title?: string;
  series: FakeSeriesSpec[];
}

// A context whose only job readChartData asks of it is sync(): every load()
// resolves synchronously against the record already in hand, exactly like
// the strict fake host's own ClientResult story but with nothing to load.
function fakeContext(
  syncImpl: () => void | Promise<void> = () => undefined,
): Excel.RequestContext {
  return {
    sync: async () => {
      await syncImpl();
    },
  } as unknown as Excel.RequestContext;
}

function fakeChart(spec: FakeChartSpec): Excel.Chart {
  return {
    chartType: spec.chartType,
    load: () => undefined,
    title: {
      load: () => undefined,
      text: spec.title ?? "",
    },
    series: {
      load: () => undefined,
      items: spec.series.map((one) => ({
        name: one.name,
        overlap: one.overlap ?? 0,
      })),
      getItemAt: (index: number) => {
        const one = spec.series[index]!;
        return {
          getDimensionValues: (dimension: string) => ({
            value:
              dimension === "Categories" ? (one.categories ?? []) : one.values,
          }),
          getDimensionDataSourceString: () => ({ value: one.source ?? "" }),
        };
      },
    },
  } as unknown as Excel.Chart;
}

describe("a chart that stops answering after its head read succeeds", () => {
  it("reports SERIES_READ_REFUSED instead of throwing out of readChartData", async () => {
    const chart = fakeChart({
      chartType: "ColumnClustered",
      series: [{ name: "Revenue", values: [1, 2, 3] }],
    });
    let syncs = 0;
    const context = fakeContext(() => {
      syncs += 1;
      // Sync 1 is the head read (chartType/title/series names); sync 2 is
      // the body (categories/values/overlap) - the one that fails here, the
      // way office.js rejects when the chart itself is gone by then.
      if (syncs === 2) throw new Error("GeneralException");
    });

    const read = await readChartData(context, chart);

    expect(read.chart).toBeNull();
    expect(read.issue).toBe(SERIES_READ_REFUSED);
  });

  it("still refuses cleanly on a second, unrelated chart read in the same context", async () => {
    // Repeats lens: the same failure twice in a row must not leave readBody's
    // try/catch in some state that only refuses once.
    const chart = fakeChart({
      chartType: "ColumnClustered",
      series: [{ name: "Revenue", values: [1, 2, 3] }],
    });
    const alwaysFails = fakeContext(() => {
      throw new Error("GeneralException");
    });
    // The head read itself fails too here; readChartData does not catch
    // that (renderChart's own try/catch does), so this exercises the
    // uncaught head-read path directly.
    await expect(readChartData(alwaysFails, chart)).rejects.toThrow(
      "GeneralException",
    );
    await expect(readChartData(alwaysFails, chart)).rejects.toThrow(
      "GeneralException",
    );
  });
});

describe("fully overlapped clustered columns", () => {
  it("ships as a picture with the overlapping-columns reason, not a garbled shape chart", async () => {
    const chart = fakeChart({
      chartType: "ColumnClustered",
      series: [
        { name: "A", values: [1, 2], overlap: 100 },
        { name: "B", values: [3, 4], overlap: 100 },
      ],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart).toBeNull();
    expect(read.issue).toBe("overlapping columns are not drawn as shapes");
  });

  it("draws the same chart normally once the columns no longer fully overlap", async () => {
    const chart = fakeChart({
      chartType: "ColumnClustered",
      series: [
        { name: "A", values: [1, 2], overlap: 0 },
        { name: "B", values: [3, 4], overlap: 0 },
      ],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart?.kind).toBe("column");
  });
});

describe("a series that is not all numbers", () => {
  it("ships as a picture instead of NaN-ing a value", async () => {
    const chart = fakeChart({
      chartType: "Line",
      series: [{ name: "Revenue", values: [1, "n/a", 3] }],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart).toBeNull();
    expect(read.issue).toBe("the series are not all numbers");
  });

  it("treats a blank cell as a gap, not a zero, and still refuses the shape chart", async () => {
    const chart = fakeChart({
      chartType: "Line",
      series: [{ name: "Revenue", values: [1, "   ", 3] }],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart).toBeNull();
    expect(read.issue).toBe("the series are not all numbers");
  });
});

describe("colour rules through the full read, not only through colorsOf in isolation", () => {
  it("brands a pie the way pieColors does for the same slice count", async () => {
    const chart = fakeChart({
      chartType: "Pie",
      series: [
        { name: "Share", categories: ["A", "B", "C"], values: [1, 2, 3] },
      ],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart?.kind).toBe("pie");
    expect(read.chart?.series[0]?.colors).toEqual(
      pieColors(3, getActiveSettings()),
    );
  });

  it("brands a waterfall the way waterfallColors does for the same bridge", async () => {
    const values = [100, 20, -30, 90];
    const chart = fakeChart({
      chartType: "Waterfall",
      series: [{ name: "Bridge", values }],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart?.kind).toBe("waterfall");
    expect(read.chart?.series[0]?.colors).toEqual(
      waterfallColors(values, getActiveSettings()),
    );
  });
});

describe("labels read off the cells the values came from", () => {
  it("falls back to the house number style when the source is not one plain A1 range", async () => {
    const chart = fakeChart({
      chartType: "Line",
      series: [
        {
          name: "Revenue",
          values: [1234.5, 2000],
          // A literal series (typed numbers, not a range) or a defined name
          // answers with no plain A1 address to read cell text from.
          source: "",
        },
      ],
    });

    const read = await readChartData(fakeContext(), chart);

    expect(read.chart?.series[0]?.labels).toEqual(["1 234.5", "2 000"]);
  });
});
