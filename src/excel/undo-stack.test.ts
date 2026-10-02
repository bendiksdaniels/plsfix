// Unit tests for the pure undo ring: depth and cell-budget eviction in
// isolation, with numbers chosen so each limit fires independently of the
// other (unlike undo.ts's real constants, where UNDO_CELL_BUDGET is exactly
// UNDO_DEPTH x the per-action cell cap and so never binds on its own).

import { describe, expect, it } from "vitest";
import { pushCapped, type Weighted } from "./undo-stack";

interface Entry extends Weighted {
  id: string;
}

function entry(id: string, cells: number): Entry {
  return { id, cells };
}

function ids(stack: Entry[]): string[] {
  return stack.map((item) => item.id);
}

describe("pushCapped", () => {
  it("keeps the newest entry at the front", () => {
    let stack: Entry[] = [];
    stack = pushCapped(stack, entry("a", 1), { maxDepth: 5, maxCells: 1000 });
    stack = pushCapped(stack, entry("b", 1), { maxDepth: 5, maxCells: 1000 });
    expect(ids(stack)).toEqual(["b", "a"]);
  });

  it("drops the oldest once the depth limit is passed", () => {
    let stack: Entry[] = [];
    for (const id of ["a", "b", "c", "d"]) {
      stack = pushCapped(stack, entry(id, 1), { maxDepth: 3, maxCells: 1000 });
    }
    expect(ids(stack)).toEqual(["d", "c", "b"]);
  });

  it("drops the oldest to fit the cell budget, well under the depth limit", () => {
    // maxDepth is generous (10): only the budget (10 cells) can be forcing
    // anything out here.
    let stack: Entry[] = [];
    stack = pushCapped(stack, entry("a", 3), { maxDepth: 10, maxCells: 10 });
    stack = pushCapped(stack, entry("b", 3), { maxDepth: 10, maxCells: 10 });
    expect(ids(stack)).toEqual(["b", "a"]); // 3 + 3 = 6, fits

    // c weighs 8: 8 + 3 + 3 = 14 > 10, so both a and b must go to fit c alone.
    stack = pushCapped(stack, entry("c", 8), { maxDepth: 10, maxCells: 10 });
    expect(ids(stack)).toEqual(["c"]);
  });

  it("drops more than one oldest entry in a single push when needed", () => {
    let stack: Entry[] = [
      entry("newer", 2),
      entry("older", 2),
      entry("oldest", 2),
    ];
    // Pushing a 5-cell entry makes the total 11, over a budget of 7; dropping
    // just "oldest" only gets to 9, so "older" has to go too, down to 7.
    stack = pushCapped(stack, entry("new", 5), { maxDepth: 10, maxCells: 7 });
    expect(ids(stack)).toEqual(["new", "newer"]);
  });

  it("never mutates the array it was given", () => {
    const before: Entry[] = [entry("a", 1)];
    const after = pushCapped(before, entry("b", 1), {
      maxDepth: 5,
      maxCells: 1000,
    });
    expect(before).toEqual([entry("a", 1)]);
    expect(after).not.toBe(before);
  });
});
