// Hunt pass 2, target 3: planBatches (src/ppt/batching.ts) as a property.
// Invariants named in plain words: every key the caller gave comes back
// exactly once, in the same order it went in; every batch stays inside the
// byte budget unless it is a single item that is over the budget on its own;
// no batch is ever left empty. fast-check generates the item lists and
// shrinks any counterexample; the seed and run count are fixed so this stays
// deterministic and fast.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { planBatches, type BatchItem } from "../../src/ppt/batching";

const SEED = 20260927;
const NUM_RUNS = 500;

// Keys are the item's own index as a string, exactly how every real caller
// (links.ts splitByBytes) builds them: unique by construction, so a batch
// can never lose or duplicate one through key collision.
const itemsArb = fc
  .array(fc.nat({ max: 1_000_000 }), { minLength: 0, maxLength: 30 })
  .map((sizes) =>
    sizes.map((bytes, index): BatchItem => ({ key: String(index), bytes })),
  );

const budgetArb = fc.integer({ min: 0, max: 1_000_000 });

describe("planBatches: every key exactly once, in order", () => {
  it("the batches concatenated equal the input keys, in the same order", () => {
    fc.assert(
      fc.property(itemsArb, budgetArb, (items, budget) => {
        const batches = planBatches(items, budget);
        const flattened = batches.flat();
        expect(flattened).toEqual(items.map((item) => item.key));
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("no batch is ever empty", () => {
    fc.assert(
      fc.property(itemsArb, budgetArb, (items, budget) => {
        const batches = planBatches(items, budget);
        expect(batches.every((batch) => batch.length > 0)).toBe(true);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("every batch stays inside the budget, unless it is a single item over budget on its own", () => {
    fc.assert(
      fc.property(itemsArb, budgetArb, (items, budget) => {
        const byKey = new Map(items.map((item) => [item.key, item.bytes]));
        const batches = planBatches(items, budget);
        for (const batch of batches) {
          const total = batch.reduce((sum, key) => sum + byKey.get(key)!, 0);
          const isLoneOversizeItem =
            batch.length === 1 && byKey.get(batch[0]!)! > budget;
          expect(total <= budget || isLoneOversizeItem).toBe(true);
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  // The shrunk counterexample this actually found once, pinned on its own: an
  // item exactly the size of the budget must NOT be forced to travel alone -
  // planBatches's own "batch.length > 0 && used + item.bytes > budgetBytes"
  // check reads "over", strictly, so a same-size neighbour packs beside it.
  it("an item exactly the size of the budget shares a batch with a zero-size neighbour", () => {
    const items: BatchItem[] = [
      { key: "0", bytes: 10 },
      { key: "1", bytes: 0 },
    ];
    expect(planBatches(items, 10)).toEqual([["0", "1"]]);
  });
});
