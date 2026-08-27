import { describe, expect, it } from "vitest";
import { bridgeSeries, cagr, formatCagrLabel } from "./chartmath";

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
