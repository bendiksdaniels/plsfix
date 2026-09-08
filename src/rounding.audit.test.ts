// The allocator on the groups a reconciliation block actually holds: blanks
// (Excel marshals an empty cell of a number range as 0), a total that rounds to
// zero, and an all-negative group. The invariant under test is always the same
// one - the parts add up to roundedTotal of the same values.
import { describe, expect, it } from "vitest";
import { allocateRounded, roundedTotal } from "./rounding";

function addsUp(values: number[], decimals: number): boolean {
  const parts = allocateRounded(values, decimals);
  const sum = parts.reduce((total, part) => total + part, 0);
  // Both sides are shifted back from whole units, so compare at the unit.
  return Math.abs(sum - roundedTotal(values, decimals)) < 10 ** -decimals / 2;
}

describe("consistent rounding over awkward groups", () => {
  it("counts a blank cell as the zero Excel hands it", () => {
    // A blank inside the group keeps its position, so every sibling cell reads
    // the same allocation off the same range.
    expect(allocateRounded([1.5, 0, 1.5], 0)).toEqual([2, 0, 1]);
    expect(addsUp([1.5, 0, 1.5], 0)).toBe(true);
  });

  it("holds a total that rounds to zero", () => {
    expect(roundedTotal([0.5, -0.5], 0)).toBe(0);
    expect(allocateRounded([0.5, -0.5], 0)).toEqual([1, -1]);
    expect(addsUp([0.5, -0.5], 0)).toBe(true);
    expect(addsUp([0.4, -0.4, 0.1], 0)).toBe(true);
  });

  it("holds an all-negative group at its own rounded total", () => {
    expect(roundedTotal([-1.5, -1.5], 0)).toBe(-3);
    expect(allocateRounded([-1.5, -1.5], 0)).toEqual([-1, -2]);
    expect(addsUp([-1.5, -1.5], 0)).toBe(true);
    expect(addsUp([-0.4, -0.4, -0.4], 0)).toBe(true);
    expect(addsUp([-1234.567, -0.004, -99.999], 2)).toBe(true);
  });

  it("holds the invariant over a spread of mixed groups", () => {
    const groups: [number[], number][] = [
      [[33.33, 33.33, 33.34], 0],
      [[0, 0, 0], 2],
      [[-2.5, 2.5], 0],
      [[1e-7, 1e-7, 1e-7], 6],
      [[1_000_000.005, -999_999.995], 2],
    ];
    for (const [values, decimals] of groups) {
      expect([values, addsUp(values, decimals)]).toEqual([values, true]);
    }
  });
});
