// The solver's own refusals and the cases a variance hunt runs into: a cell
// the sheet answered with an error value, a tolerance wider than the target,
// and a target only the empty subset would reach.
import { describe, expect, it } from "vitest";
import { RECONCILE_MAX_VALUES, solveReconciliation } from "./reconcile";

describe("the reconciliation solver at its edges", () => {
  it("refuses a value that is not finite", () => {
    // #REF! and #N/A never reach it - the adapter keeps numbers only - but a
    // caller that skipped that filter gets told, not a silent NaN answer.
    expect(() => solveReconciliation([1, Number.NaN, 3], 4, 0)).toThrow(
      "Every value must be finite.",
    );
    expect(() =>
      solveReconciliation([1, Number.POSITIVE_INFINITY], 1, 0),
    ).toThrow("Every value must be finite.");
  });

  it("never answers with no cells at all", () => {
    // Zero is the empty subset's sum, and "no cells" is not an answer to a
    // variance hunt: the closest real combination is.
    const match = solveReconciliation([5, -5, 12], 0, 1);
    expect(match?.indices.length).toBeGreaterThan(0);
    expect(match?.sum).toBe(0);
  });

  it("takes the nearest sum when the tolerance swallows the target", () => {
    const match = solveReconciliation([40, 90], 10, 1000);
    expect(match).toMatchObject({ indices: [0], sum: 40, difference: 30 });
  });

  it("refuses more values than the cap", () => {
    const values = Array.from(
      { length: RECONCILE_MAX_VALUES + 1 },
      (_unused, index) => index + 1,
    );
    expect(() => solveReconciliation(values, 10, 0)).toThrow(
      "Reconciliation supports up to 34 numeric cells.",
    );
  });
});
