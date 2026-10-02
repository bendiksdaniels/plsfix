import { describe, expect, it } from "vitest";
import {
  bridgeSeries,
  cagr,
  footballField,
  type FootballRow,
  formatCagrLabel,
  seriesSpan,
  type TornadoDriver,
  tornadoSeries,
} from "./chartmath";

describe("seriesSpan", () => {
  it("drops the blank cells at both ends", () => {
    expect(seriesSpan(["", null, 100, 110, "", null])).toEqual({
      first: 2,
      last: 3,
    });
  });

  it("keeps blanks inside the line, and text at its ends", () => {
    expect(seriesSpan([100, "", 121])).toEqual({ first: 0, last: 2 });
    expect(seriesSpan(["n/a", 121])).toEqual({ first: 0, last: 1 });
  });

  it("answers null for a line holding nothing", () => {
    expect(seriesSpan(["", null])).toBeNull();
    expect(seriesSpan([])).toBeNull();
  });
});

describe("bridgeSeries", () => {
  it("splits a bridge into base, rise and fall", () => {
    expect(bridgeSeries([100, 20, -30, 90])).toEqual({
      base: [0, 100, 90, 0],
      rise: [100, 20, 0, 90],
      fall: [0, 0, 30, 0],
    });
  });

  it("keeps the running level in base + rise", () => {
    const { base, rise } = bridgeSeries([100, 20, -30, 90]);
    expect(base.map((value, index) => value + (rise[index] ?? 0))).toEqual([
      100, 120, 90, 90,
    ]);
  });

  it("stacks all-positive deltas on a growing base", () => {
    expect(bridgeSeries([100, 20, 30, 150])).toEqual({
      base: [0, 100, 120, 0],
      rise: [100, 20, 30, 150],
      fall: [0, 0, 0, 0],
    });
  });

  it("drops the base with all-negative deltas", () => {
    expect(bridgeSeries([100, -20, -30, 50])).toEqual({
      base: [0, 80, 50, 0],
      rise: [100, 0, 0, 50],
      fall: [0, 20, 30, 0],
    });
  });

  it("carries the sign of a negative closing total", () => {
    expect(bridgeSeries([100, -160, -60])).toEqual({
      base: [0, -60, 0],
      rise: [100, 0, -60],
      fall: [0, 160, 0],
    });
  });

  it("leaves a zero delta flat", () => {
    expect(bridgeSeries([100, 0, 100])).toEqual({
      base: [0, 100, 0],
      rise: [100, 0, 100],
      fall: [0, 0, 0],
    });
  });

  it("rejects anything shorter than two totals and a delta", () => {
    expect(() => bridgeSeries([])).toThrow();
    expect(() => bridgeSeries([100])).toThrow();
    expect(() => bridgeSeries([100, 90])).toThrow();
  });
});

describe("cagr", () => {
  it("compounds the growth over the periods", () => {
    expect(cagr(100, 121, 2)).toBeCloseTo(0.1, 10);
    expect(cagr(100, 90, 1)).toBeCloseTo(-0.1, 10);
  });

  it("rejects a period count below one", () => {
    expect(() => cagr(100, 121, 0)).toThrow();
    expect(() => cagr(100, 121, -2)).toThrow();
  });

  it("rejects values a compound rate cannot describe", () => {
    expect(() => cagr(0, 121, 2)).toThrow();
    expect(() => cagr(-100, 121, 2)).toThrow();
    expect(() => cagr(100, -121, 2)).toThrow();
  });
});

describe("formatCagrLabel", () => {
  it("always shows the sign at one decimal", () => {
    expect(formatCagrLabel(0.1234)).toBe("CAGR +12.3%");
    expect(formatCagrLabel(-0.0412)).toBe("CAGR -4.1%");
    expect(formatCagrLabel(0)).toBe("CAGR +0.0%");
  });

  it("never prints a negative zero", () => {
    expect(formatCagrLabel(-0.0004)).toBe("CAGR +0.0%");
  });
});

describe("tornadoSeries", () => {
  const drivers: TornadoDriver[] = [
    { label: "Volume", low: 90, high: 115 },
    { label: "Price", low: 60, high: 140 },
    { label: "Mix", low: 95, high: 105 },
  ];

  it("ranks drivers by swing and reports deltas from the stated base", () => {
    expect(tornadoSeries(drivers, 100)).toEqual({
      labels: ["Price", "Volume", "Mix"],
      low: [-40, -10, -5],
      high: [40, 15, 5],
      base: 100,
    });
  });

  it("falls back to the mean of every low and high", () => {
    const { base, low, high } = tornadoSeries(
      [
        { label: "A", low: 80, high: 120 },
        { label: "B", low: 90, high: 110 },
      ],
      null,
    );
    expect(base).toBe(100);
    expect(low).toEqual([-20, -10]);
    expect(high).toEqual([20, 10]);
  });

  it("keeps tied drivers in the order they were selected in", () => {
    expect(
      tornadoSeries(
        [
          { label: "First", low: 95, high: 105 },
          { label: "Second", low: 90, high: 100 },
          { label: "Third", low: 0, high: 10 },
        ],
        100,
      ).labels,
    ).toEqual(["First", "Second", "Third"]);
  });

  it("ranks an inverted driver by the size of its swing, not its sign", () => {
    expect(
      tornadoSeries(
        [
          { label: "Normal", low: 95, high: 105 },
          { label: "Inverted", low: 130, high: 70 },
        ],
        100,
      ),
    ).toEqual({
      labels: ["Inverted", "Normal"],
      low: [30, -5],
      high: [-30, 5],
      base: 100,
    });
  });

  it("leaves a driver that moves nothing flat at the base", () => {
    const { low, high } = tornadoSeries(
      [
        { label: "Flat", low: 100, high: 100 },
        { label: "Live", low: 80, high: 120 },
      ],
      100,
    );
    expect(low).toEqual([-20, 0]);
    expect(high).toEqual([20, 0]);
  });

  it("keeps a negative base and its deltas", () => {
    expect(tornadoSeries(drivers, -50).low).toEqual([110, 140, 145]);
  });

  it("refuses a chart nothing can be ranked in", () => {
    expect(() => tornadoSeries([], null)).toThrow(
      "tornado: need at least two drivers",
    );
    expect(() => tornadoSeries([drivers[0]!], null)).toThrow(
      "tornado: need at least two drivers",
    );
  });

  it("refuses outcomes and bases that are not finite numbers", () => {
    const broken = [{ label: "A", low: Number.NaN, high: 1 }, drivers[0]!];
    expect(() => tornadoSeries(broken, null)).toThrow(
      "tornado: every low and high must be a finite number",
    );
    expect(() => tornadoSeries(drivers, Number.POSITIVE_INFINITY)).toThrow(
      "tornado: the base must be a finite number",
    );
  });
});

describe("footballField", () => {
  const rows: FootballRow[] = [
    { label: "DCF", low: 90, high: 130 },
    { label: "Trading comps", low: 100, high: 120 },
  ];

  it("splits each range into its floor and the band above it", () => {
    expect(footballField(rows)).toEqual({
      labels: ["DCF", "Trading comps"],
      low: [90, 100],
      high: [130, 120],
      range: [40, 20],
      swapped: [false, false],
      swaps: 0,
    });
  });

  it("swaps a row entered high first and counts it", () => {
    const field = footballField([
      { label: "DCF", low: 130, high: 90 },
      rows[1]!,
    ]);
    expect(field.low).toEqual([90, 100]);
    expect(field.high).toEqual([130, 120]);
    expect(field.range).toEqual([40, 20]);
    expect(field.swapped).toEqual([true, false]);
    expect(field.swaps).toBe(1);
  });

  it("keeps a row whose low and high are the same, as a zero band", () => {
    const field = footballField([
      { label: "52-week", low: 8, high: 8 },
      rows[0]!,
    ]);
    expect(field.range).toEqual([0, 40]);
    expect(field.swaps).toBe(0);
  });

  it("keeps the rows in the order they were selected", () => {
    expect(footballField([rows[1]!, rows[0]!]).labels).toEqual([
      "Trading comps",
      "DCF",
    ]);
  });

  it("refuses fewer than two rows", () => {
    expect(() => footballField([rows[0]!])).toThrow(
      "football field: need at least two rows",
    );
  });

  it("refuses a low or a high that is not a finite number", () => {
    expect(() =>
      footballField([{ label: "DCF", low: Number.NaN, high: 1 }, rows[0]!]),
    ).toThrow("football field: every low and high must be a finite number");
  });
});
