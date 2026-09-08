// Audit coverage for object-math.ts's defensive guards: object-tools.ts
// always validates the selection count through requireCount before calling
// into these pure functions, so their own too-few-boxes branches are never
// reached from the real call path and object-math.test.ts never exercises
// them either. Proves each guard is a safe no-op, not a crash, and answers
// two of slice H's seeds directly (distribute with exactly two shapes, swap
// of unequal sizes) without needing the office.js layer at all.

import { describe, expect, it } from "vitest";
import {
  alignBoxes,
  distributeBoxes,
  matchSize,
  swapBoxes,
  type ObjectBox,
} from "./object-math";

const one: ObjectBox[] = [{ id: "a", left: 0, top: 0, width: 10, height: 10 }];
const two: ObjectBox[] = [
  ...one,
  { id: "b", left: 100, top: 100, width: 20, height: 20 },
];

describe("object-math guard clauses (unreachable via object-tools.ts, defensive)", () => {
  it("alignBoxes is a no-op on an empty selection or a single object", () => {
    expect(alignBoxes([], "left")).toEqual([]);
    expect(alignBoxes(one, "left")).toEqual([]);
  });

  it("distributeBoxes is a no-op below three objects, exactly two included", () => {
    // The seed's question - "distribute with exactly two shapes: a no-op or
    // a sentence?" - is answered a layer up: distributeSelected() refuses two
    // objects with "Select at least three objects to distribute." before this
    // function ever runs. Here, at the pure-math layer, two boxes is a no-op.
    expect(distributeBoxes([], "horizontal")).toEqual([]);
    expect(distributeBoxes(one, "horizontal")).toEqual([]);
    expect(distributeBoxes(two, "horizontal")).toEqual([]);
    expect(distributeBoxes(two, "vertical")).toEqual([]);
  });

  it("matchSize is a no-op on an empty selection or a single object", () => {
    expect(matchSize([])).toEqual([]);
    expect(matchSize(one)).toEqual([]);
  });

  it("swapBoxes is a no-op on anything but exactly two objects", () => {
    const three: ObjectBox[] = [
      ...two,
      { id: "c", left: 200, top: 200, width: 5, height: 5 },
    ];
    expect(swapBoxes([])).toEqual([]);
    expect(swapBoxes(one)).toEqual([]);
    expect(swapBoxes(three)).toEqual([]);
  });

  it("swaps top-left corners, not centres, keeping each box's own size", () => {
    const small: ObjectBox = {
      id: "small",
      left: 900,
      top: 500,
      width: 20,
      height: 20,
    };
    const big: ObjectBox = {
      id: "big",
      left: 10,
      top: 10,
      width: 200,
      height: 200,
    };
    const moves = swapBoxes([small, big]);
    expect(moves).toEqual([
      { id: "small", left: 10, top: 10 },
      { id: "big", left: 900, top: 500 },
    ]);
    // Documents the consequence: the big box's far edge now runs past a
    // standard 960x540 slide, because a corner swap carries no size and no
    // clamp - the same "no clamping" rule align and distribute already
    // follow against the selection's own bounding box.
    const bigMove = moves.find((move) => move.id === "big")!;
    const bigLeft = bigMove.left ?? big.left;
    const bigTop = bigMove.top ?? big.top;
    expect(bigLeft + big.width).toBeGreaterThan(960);
    expect(bigTop + big.height).toBeGreaterThan(540);
  });
});
