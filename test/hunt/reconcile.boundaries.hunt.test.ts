// Find a combination at the RECONCILE_MAX_VALUES edge. The 35-value refusal
// is pinned already (src/reconcile.audit.test.ts, test/reconcile.integration.
// test.ts); the 34-value success side of that same boundary was not.

import { describe, expect, it } from "vitest";
import { RECONCILE_MAX_VALUES, solveReconciliation } from "../../src/reconcile";

describe("solveReconciliation at the value-count cap", () => {
  it("still solves at exactly the cap instead of refusing one value early", () => {
    // 34 values of 1, target 34: only the full set reaches it exactly.
    const values = Array.from({ length: RECONCILE_MAX_VALUES }, () => 1);
    const result = solveReconciliation(values, RECONCILE_MAX_VALUES, 0);
    expect(result?.indices).toHaveLength(RECONCILE_MAX_VALUES);
    expect(result?.sum).toBe(RECONCILE_MAX_VALUES);
    expect(result?.difference).toBe(0);
  });

  it("finds the minimal subset at the cap when a smaller one exists", () => {
    // 33 filler values that cannot reach 100 alone (they are 1 each) plus one
    // cell worth exactly the target: the single cell must still win at 34
    // values, the same "fewest cells" rule proven at smaller sizes.
    const values = [
      100,
      ...Array.from({ length: RECONCILE_MAX_VALUES - 1 }, () => 1),
    ];
    const result = solveReconciliation(values, 100, 0);
    expect(result?.indices).toEqual([0]);
  });
});
