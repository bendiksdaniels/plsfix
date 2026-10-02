// Pass-1 hunt: the largest-remainder allocator (src/rounding.ts) at its own
// edges, directly (callers own the cell cap): tie stability under reordering,
// ten thousand parts, the decimals boundary, negative zero, float artefacts on
// the group total, and a huge magnitude spread.
import { describe, expect, it } from "vitest";
import {
  allocateRounded,
  ROUNDING_MAX_DECIMALS,
  roundedTotal,
} from "../../src/rounding";

function addsUp(values: number[], decimals: number): boolean {
  const parts = allocateRounded(values, decimals);
  const sum = parts.reduce((total, part) => total + part, 0);
  return Math.abs(sum - roundedTotal(values, decimals)) < 10 ** -decimals / 2;
}

describe("negative zero never comes back as -0", () => {
  it("normalizes a total that rounds to zero from the negative side", () => {
    expect(Object.is(roundedTotal([-0.3], 0), 0)).toBe(true);
    // -0.2 + -0.1 sums to a magnitude under one half, not a tie, so this
    // stays on the "rounds to zero" side rather than "half away from zero".
    expect(Object.is(roundedTotal([-0.2, -0.1], 0), 0)).toBe(true);
    expect(Object.is(roundedTotal([-1, 1], 0), 0)).toBe(true);
    expect(Object.is(roundedTotal([-0.0001], 2), 0)).toBe(true);
  });

  it("normalizes every allocated part, not only the total", () => {
    const parts = allocateRounded([-0.1, -0.2], 0);
    expect(parts).toEqual([0, 0]);
    for (const part of parts) expect(Object.is(part, -0)).toBe(false);
  });

  it("holds at negative decimal precisions too", () => {
    expect(Object.is(roundedTotal([-40], -2), 0)).toBe(true);
    expect(Object.is(allocateRounded([-40], -2)[0], -0)).toBe(false);
  });
});

describe("ties are stable and follow position, not value", () => {
  it("repeats the same allocation for the same order", () => {
    const values = [1.5, 1.5, 1.5, 1.5];
    const first = allocateRounded(values, 0);
    expect(first).toEqual([2, 2, 1, 1]);
    expect(allocateRounded([...values], 0)).toEqual(first);
  });

  it("moves the bump to the new position when the tied values swap places", () => {
    // Same two values, same tied 0.5 remainder each; only their order in the
    // range differs. The extra unit follows the position, not the number.
    expect(allocateRounded([1.5, 3.5], 0)).toEqual([2, 3]);
    expect(allocateRounded([3.5, 1.5], 0)).toEqual([4, 1]);
  });
});

describe("scale: ten thousand parts", () => {
  it("holds the total and the one-unit bound at ten thousand parts", () => {
    let seed = 11;
    const values = Array.from({ length: 10_000 }, () => {
      seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
      return (seed / 2_147_483_648) * 2000 - 1000;
    });

    for (const decimals of [-2, 0, 2, 5]) {
      const parts = allocateRounded(values, decimals);
      expect(parts.length).toBe(values.length);
      expect(addsUp(values, decimals)).toBe(true);

      const unit = 10 ** -decimals;
      parts.forEach((part, index) => {
        expect(Math.abs(part - values[index]!)).toBeLessThan(unit);
      });
    }
  });

  it("still breaks ties by position at ten thousand equal parts", () => {
    const values = Array.from({ length: 10_000 }, () => 1 / 3);
    const parts = allocateRounded(values, 0);
    // 10,000/3 = 3333.33...; the total rounds to 3333, one short of the 3,334
    // whole units the parts alone would floor to zero and lose entirely.
    expect(roundedTotal(values, 0)).toBe(3333);
    expect(parts.reduce((sum, part) => sum + part, 0)).toBe(3333);
    // Every part is 0 or 1, and the 1s are the earliest positions.
    const ones = parts.filter((part) => part === 1).length;
    expect(new Set(parts)).toEqual(new Set([0, 1]));
    expect(parts.slice(0, ones).every((part) => part === 1)).toBe(true);
    expect(parts.slice(ones).every((part) => part === 0)).toBe(true);
  });
});

describe("the decimals boundary", () => {
  it("allows exactly the max decimals in both directions", () => {
    expect(() => allocateRounded([1, 2], ROUNDING_MAX_DECIMALS)).not.toThrow();
    expect(() => allocateRounded([1, 2], -ROUNDING_MAX_DECIMALS)).not.toThrow();
    expect(() => roundedTotal([1, 2], ROUNDING_MAX_DECIMALS)).not.toThrow();
    expect(() => roundedTotal([1, 2], -ROUNDING_MAX_DECIMALS)).not.toThrow();
  });

  it("refuses one past the max decimals on the negative side too", () => {
    expect(() => allocateRounded([1, 2], -ROUNDING_MAX_DECIMALS - 1)).toThrow(
      "rounding: decimals must be a whole number",
    );
    expect(() => roundedTotal([1, 2], -ROUNDING_MAX_DECIMALS - 1)).toThrow(
      "rounding: decimals must be a whole number",
    );
  });
});

describe("float artefacts named by the brief", () => {
  it("keeps 0.1 + 0.2 exact at the decimal the group total rounds to", () => {
    expect(roundedTotal([0.1, 0.2], 1)).toBe(0.3);
    expect(allocateRounded([0.1, 0.2], 1)).toEqual([0.1, 0.2]);
    expect(addsUp([0.1, 0.2], 1)).toBe(true);
  });

  it("keeps a 1.005-shaped group total exact, not the binary-rounded 100.49...", () => {
    expect(roundedTotal([1.005, 1.005], 2)).toBe(2.01);
    expect(allocateRounded([1.005, 1.005], 2)).toEqual([1.01, 1]);
  });
});

describe("a huge magnitude spread of both signs", () => {
  it("keeps every part within one unit of its floor", () => {
    let seed = 5;
    const values = Array.from({ length: 500 }, (_unused, index) => {
      seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
      const sign = seed % 2 === 0 ? 1 : -1;
      // Every third value is tiny, the rest span up to 100 million: the kind
      // of spread that could push a floor-sum and a float total apart.
      const magnitude = index % 3 === 0 ? 1e-4 : 1e8 * (seed / 2_147_483_648);
      return sign * magnitude;
    });

    for (const decimals of [-4, 0, 3]) {
      expect(() => allocateRounded(values, decimals)).not.toThrow();
      const parts = allocateRounded(values, decimals);
      // No part may differ from its floor by more than one unit either way.
      const unit = 10 ** -decimals;
      const units = values.map((value) => value / unit);
      parts.forEach((part, index) => {
        const floor = Math.floor(units[index]!) * unit;
        const diff = Math.round((part - floor) / unit);
        expect([0, 1]).toContain(diff);
      });
    }
  });
});

describe("all-equal, one part, and zero-remainder ties", () => {
  it("leaves an all-equal whole group alone at any size", () => {
    expect(
      allocateRounded(
        Array.from({ length: 50 }, () => 7),
        0,
      ),
    ).toEqual(Array.from({ length: 50 }, () => 7));
  });

  it("hands the one part its own rounded value at every decimals sign", () => {
    for (const decimals of [-3, 0, 3]) {
      expect(allocateRounded([1234.5678], decimals)).toEqual([
        roundedTotal([1234.5678], decimals),
      ]);
    }
    expect(allocateRounded([-1234.5678], 2)).toEqual([
      roundedTotal([-1234.5678], 2),
    ]);
  });

  it("bumps nobody when every remainder is already zero", () => {
    expect(allocateRounded([10, 0, -10, 5], 0)).toEqual([10, 0, -10, 5]);
    expect(roundedTotal([10, 0, -10, 5], 0)).toBe(5);
  });
});
