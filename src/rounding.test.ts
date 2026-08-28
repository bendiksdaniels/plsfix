// The rounding contract: the parts add up to the total at the same precision,
// every part stays within one unit of its value, and the same inputs always
// produce the same allocation - the reason a hundred sibling cells can each
// recompute the group on their own and still agree.
import { describe, expect, it } from "vitest";
import {
  allocateRounded,
  ROUNDING_MAX_DECIMALS,
  roundedTotal,
} from "./rounding";

// Adding the parts back up in the same precision the caller asked for: the
// invariant the whole feature exists to keep.
function partsTotal(parts: number[], decimals: number): number {
  return roundedTotal(parts, decimals);
}

describe("allocateRounded", () => {
  it("splits 100 across three thirds without losing the unit", () => {
    const values = [33.333, 33.333, 33.334];
    expect(allocateRounded(values, 0)).toEqual([33, 33, 34]);
    expect(partsTotal(allocateRounded(values, 0), 0)).toBe(
      roundedTotal(values, 0),
    );
    expect(roundedTotal(values, 0)).toBe(100);
  });

  it("breaks a tie by position, not by value", () => {
    const third = 10 / 3;
    expect(allocateRounded([third, third, third], 0)).toEqual([4, 3, 3]);
    expect(allocateRounded([third, third, third], 1)).toEqual([3.4, 3.3, 3.3]);
  });

  it("mirrors the split for negative values", () => {
    const values = [-33.333, -33.333, -33.334];
    expect(allocateRounded(values, 0)).toEqual([-33, -33, -34]);
    expect(roundedTotal(values, 0)).toBe(-100);
  });

  it("allocates across mixed signs", () => {
    const values = [2.5, -1.2, -0.4];
    expect(allocateRounded(values, 0)).toEqual([2, -1, 0]);
    expect(roundedTotal(values, 0)).toBe(1);
  });

  it("leaves zeros and whole numbers alone", () => {
    expect(allocateRounded([0, 0, 0], 0)).toEqual([0, 0, 0]);
    expect(allocateRounded([10, 20, 30], 2)).toEqual([10, 20, 30]);
    expect(allocateRounded([], 0)).toEqual([]);
  });

  it("rounds a single value half away from zero", () => {
    expect(allocateRounded([2.5], 0)).toEqual([3]);
    expect(allocateRounded([-2.5], 0)).toEqual([-3]);
    expect(roundedTotal([-0.5], 0)).toBe(-1);
    expect(roundedTotal([0.5], 0)).toBe(1);
  });

  it("keeps decimals exact rather than binary-approximate", () => {
    // 1.005 * 100 is 100.49999999999999 in binary; a multiply here would floor
    // to 100 and hand the unit to the wrong cell.
    const values = [1.005, 2.005, 3.005];
    expect(allocateRounded(values, 2)).toEqual([1.01, 2.01, 3]);
    expect(roundedTotal(values, 2)).toBe(6.02);
  });

  it("rounds to thousands on negative decimals", () => {
    expect(allocateRounded([1400, 1400, 1400], -3)).toEqual([2000, 1000, 1000]);
    expect(roundedTotal([1400, 1400, 1400], -3)).toBe(4000);
  });

  // A hand-checked block: the parts a modeller would read off the sheet.
  it("keeps a percentage split adding to the whole", () => {
    const values = [0.31428, 0.31428, 0.37144];
    expect(allocateRounded(values, 3)).toEqual([0.314, 0.314, 0.372]);
    expect(roundedTotal(values, 3)).toBe(1);
  });

  it("holds the total and the one-unit bound over a long series", () => {
    // Deterministic pseudo-random values, so a failure is reproducible.
    let seed = 7;
    const values = Array.from({ length: 200 }, () => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return (seed / 2147483648) * 2000 - 500;
    });

    for (const decimals of [-1, 0, 1, 2]) {
      const parts = allocateRounded(values, decimals);
      expect(partsTotal(parts, decimals)).toBe(roundedTotal(values, decimals));
      const unit = 10 ** -decimals;
      parts.forEach((part, index) => {
        expect(Math.abs(part - values[index]!)).toBeLessThan(unit);
      });
    }
  });

  it("gives every caller the same allocation", () => {
    const values = [1.4, 1.4, 1.4, 1.4, 1.4];
    const first = allocateRounded(values, 0);
    expect(allocateRounded([...values], 0)).toEqual(first);
    expect(first).toEqual([2, 2, 1, 1, 1]);
  });

  it("refuses values and precisions it cannot allocate", () => {
    expect(() => allocateRounded([1, Number.NaN], 0)).toThrow(
      "rounding: every value must be a finite number",
    );
    expect(() => allocateRounded([1, Number.POSITIVE_INFINITY], 0)).toThrow(
      "rounding: every value must be a finite number",
    );
    expect(() => allocateRounded([1, 2], 1.5)).toThrow(
      "rounding: decimals must be a whole number",
    );
    expect(() => allocateRounded([1, 2], ROUNDING_MAX_DECIMALS + 1)).toThrow(
      "rounding: decimals must be a whole number",
    );
    expect(() => roundedTotal([1, 2], Number.NaN)).toThrow(
      "rounding: decimals must be a whole number",
    );
  });
});
