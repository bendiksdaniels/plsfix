// Find a combination: a zero-summing subset confined to the meet-in-the-middle
// solver's second (deduplicated) half. Attacks the boundary between the
// "fewest cells" tie-break and the empty-selection guard in solveReconciliation.

import { describe, expect, it } from "vitest";
import { solveReconciliation } from "../../src/reconcile";

describe("solveReconciliation: a zero-sum subset in the second half", () => {
  it("prefers the single zero cell over a pair that also nets to zero", () => {
    // cut = floor(3/2) = 1: left = [7], right = [-7, 0]. The lone "0" sits in
    // the deduplicated half, tied on sum with that half's empty pick.
    const result = solveReconciliation([7, -7, 0], 0, 0);
    expect(result?.indices).toEqual([2]);
    expect(result?.sum).toBe(0);
  });

  it("still finds the single zero cell with a second zero-valued cell present", () => {
    const result = solveReconciliation([7, -7, 0, 0], 0, 0);
    expect(result?.indices).toHaveLength(1);
    expect(result?.sum).toBe(0);
  });

  it("finds an offsetting pair confined entirely to the second half", () => {
    // cut = floor(3/2) = 1: left = [3], right = [5, -5]. Nothing on the left
    // contributes; the answer is entirely the deduplicated half's own pair.
    // Before the fix this came back null: the pair's zero sum collapsed into
    // the half's empty pick, and the "both sides empty" guard then refused to
    // pair it with the other half's empty pick too.
    const result = solveReconciliation([3, 5, -5], 0, 0);
    expect(result).not.toBeNull();
    expect(result?.indices).toEqual([1, 2]);
    expect(result?.sum).toBe(0);
  });

  it("still refuses the truly empty selection for a target of zero", () => {
    // No cell is zero and nothing offsets: the only exact match is selecting
    // nothing at all, which must stay refused rather than reported as "0 cells".
    expect(solveReconciliation([4, 9, 15], 0, 0)).toBeNull();
  });

  it("keeps the existing fewest-cells tie-break for a non-zero sum", () => {
    // Unaffected control case: three right-half combinations reach 10 with no
    // zero-sum subset involved at all.
    expect(
      solveReconciliation([1, 9, 2, 8, 3, 7, 4, 6, 10], 10, 0)?.indices,
    ).toEqual([8]);
  });
});
