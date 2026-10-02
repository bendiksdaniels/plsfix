import { describe, expect, it } from "vitest";
import {
  RECONCILE_MAX_VALUES,
  reconcileSummary,
  solveReconciliation,
} from "./reconcile";

describe("variance reconciliation", () => {
  it("finds a subset that equals the target across both halves", () => {
    expect(solveReconciliation([125, -40, 17, 300, -85, 9], 49, 0.001)).toEqual(
      {
        indices: [0, 4, 5],
        sum: 49,
        difference: 0,
      },
    );
  });

  it("chooses the smallest exact combination and supports a tolerance", () => {
    expect(solveReconciliation([5, 5, 10], 10, 0)?.indices).toEqual([2]);
    // Three right-half combinations reach 10 (3+7, 4+6, 10): the single cell
    // must win even when it sorts behind the pairs.
    expect(
      solveReconciliation([1, 9, 2, 8, 3, 7, 4, 6, 10], 10, 0)?.indices,
    ).toEqual([8]);
    const close = solveReconciliation([33.33, 66.66], 100, 0.02);
    expect(close?.indices).toEqual([0, 1]);
    expect(close?.difference).toBeCloseTo(-0.01);
  });

  it("returns null when nothing reconciles and rejects unsafe inputs", () => {
    expect(solveReconciliation([1, 2, 4], 20, 0)).toBeNull();
    expect(() =>
      solveReconciliation(Array(RECONCILE_MAX_VALUES + 1).fill(1), 2, 0),
    ).toThrow("up to");
    expect(() => solveReconciliation([1], 1, -1)).toThrow("Tolerance");
  });
});

describe("reconcileSummary", () => {
  it("groups thousands by a space in Latvian and drops a noise-level difference", () => {
    expect(
      reconcileSummary({ count: 3, sum: 1234.56, difference: 1e-12 }, "lv"),
    ).toBe("3 cells selected · Sum 1 234.56 · Variance 0");
  });

  it("groups thousands by a comma in English and keeps a real negative variance", () => {
    expect(
      reconcileSummary({ count: 5, sum: 98765, difference: -12.34 }, "en"),
    ).toBe("5 cells selected · Sum 98,765 · Variance -12.34");
  });
});
