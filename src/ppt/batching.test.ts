import { describe, expect, it } from "vitest";
import { planBatches, REPAINT_BUDGET_BYTES, type BatchItem } from "./batching";

function items(...bytes: number[]): BatchItem[] {
  return bytes.map((size, index) => ({ key: String(index), bytes: size }));
}

describe("planBatches", () => {
  it("fills a batch to the budget and starts the next one", () => {
    expect(planBatches(items(4, 4, 4, 4, 4), 10)).toEqual([
      ["0", "1"],
      ["2", "3"],
      ["4"],
    ]);
  });

  it("packs an exact fit into one batch", () => {
    expect(planBatches(items(6, 4), 10)).toEqual([["0", "1"]]);
    expect(planBatches(items(6, 5), 10)).toEqual([["0"], ["1"]]);
  });

  // A picture nothing can pack beside it still has to be painted, so it travels
  // on its own rather than blocking the run or riding with a neighbour.
  it("gives an item past the budget a batch of its own", () => {
    expect(planBatches(items(1, 99, 1), 10)).toEqual([["0"], ["1"], ["2"]]);
    expect(planBatches(items(99, 99), 10)).toEqual([["0"], ["1"]]);
  });

  it("keeps the order it was given and loses nothing", () => {
    const plan = planBatches(items(3, 3, 3, 3, 3, 3, 3), 7);
    expect(plan.flat()).toEqual(["0", "1", "2", "3", "4", "5", "6"]);
    expect(plan).toEqual(planBatches(items(3, 3, 3, 3, 3, 3, 3), 7));
  });

  it("has nothing to plan for an empty list", () => {
    expect(planBatches([], REPAINT_BUDGET_BYTES)).toEqual([]);
  });

  // Weightless rows (a payload that is somehow empty) must not spin: they land
  // in one batch, never in an endless flush of empty ones.
  it("keeps zero-byte items in one batch", () => {
    expect(planBatches(items(0, 0, 0), 0)).toEqual([["0", "1", "2"]]);
  });

  // The real budget: sixty small pictures are one round trip, three full-slide
  // renders are three.
  it("batches a deck at the repaint budget", () => {
    const small = new Array<number>(60).fill(1024);
    expect(planBatches(items(...small), REPAINT_BUDGET_BYTES)).toHaveLength(1);
    const big = new Array<number>(3).fill(5 * 1024 * 1024);
    expect(planBatches(items(...big), REPAINT_BUDGET_BYTES)).toHaveLength(3);
  });
});
