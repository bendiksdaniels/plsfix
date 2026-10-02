// Hunt pass 2, target 3: object-math.ts (alignBoxes, distributeBoxes,
// matchSize, swapBoxes) as properties over random selections. Plain-word
// invariants: align collapses every box onto one shared line on the aligned
// edge; distribute keeps the two outer edges fixed and only moves the inner
// ones; matchSize copies the reference's own size onto every target and
// never touches the reference itself; swap trades positions only, sizes stay
// put; none of the four ever emits a negative width or height. fast-check
// generates the selections; seed and run count are fixed for speed.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  alignBoxes,
  distributeBoxes,
  matchSize,
  swapBoxes,
  type ObjectBox,
} from "../../src/ppt/object-math";

const SEED = 20260927;
const NUM_RUNS = 300;
const EPS = 1e-9;

// Office.js geometry: left/top can be anywhere a shape was dragged
// (off-slide included), width/height are never negative (PowerPoint itself
// refuses that on the way in - size-guard.ts).
const geometryArb = fc.record({
  left: fc.integer({ min: -500, max: 1500 }),
  top: fc.integer({ min: -500, max: 1500 }),
  width: fc.integer({ min: 0, max: 900 }),
  height: fc.integer({ min: 0, max: 600 }),
});

function boxesArb(min: number, max: number): fc.Arbitrary<ObjectBox[]> {
  return fc
    .array(geometryArb, { minLength: min, maxLength: max })
    .map((list) =>
      list.map((g, i): ObjectBox => ({ id: `s${String(i)}`, ...g })),
    );
}

function byId(boxes: ObjectBox[]): Map<string, ObjectBox> {
  return new Map(boxes.map((box) => [box.id, box]));
}

describe("alignBoxes: every box lands on one shared line, sizes untouched", () => {
  it.each(["left", "right", "top", "bottom", "center", "middle"] as const)(
    "%s collapses the whole selection onto the group's own edge or centre line",
    (mode) => {
      fc.assert(
        fc.property(boxesArb(2, 8), (boxes) => {
          const moves = alignBoxes(boxes, mode);
          expect(moves).toHaveLength(boxes.length);
          const source = byId(boxes);
          const left = Math.min(...boxes.map((b) => b.left));
          const right = Math.max(...boxes.map((b) => b.left + b.width));
          const top = Math.min(...boxes.map((b) => b.top));
          const bottom = Math.max(...boxes.map((b) => b.top + b.height));
          for (const move of moves) {
            const original = source.get(move.id)!;
            // Align never resizes: only left or only top ever appears.
            expect(move.width).toBeUndefined();
            expect(move.height).toBeUndefined();
            if (mode === "left") expect(move.left).toBeCloseTo(left, 6);
            if (mode === "right") {
              expect(move.left! + original.width).toBeCloseTo(right, 6);
            }
            if (mode === "center") {
              expect(move.left! + original.width / 2).toBeCloseTo(
                (left + right) / 2,
                6,
              );
            }
            if (mode === "top") expect(move.top).toBeCloseTo(top, 6);
            if (mode === "bottom") {
              expect(move.top! + original.height).toBeCloseTo(bottom, 6);
            }
            if (mode === "middle") {
              expect(move.top! + original.height / 2).toBeCloseTo(
                (top + bottom) / 2,
                6,
              );
            }
          }
        }),
        { seed: SEED, numRuns: NUM_RUNS },
      );
    },
  );

  it("fewer than two boxes always answers no moves at all", () => {
    fc.assert(
      fc.property(
        boxesArb(0, 1),
        fc.constantFrom("left", "center", "top"),
        (boxes, mode) => {
          expect(alignBoxes(boxes, mode)).toEqual([]);
        },
      ),
      { seed: SEED, numRuns: 50 },
    );
  });
});

describe("distributeBoxes: the two outer edges never move, sizes untouched", () => {
  it.each(["horizontal", "vertical"] as const)(
    "%s keeps the first box's leading edge and the last box's trailing edge fixed",
    (axis) => {
      fc.assert(
        fc.property(boxesArb(3, 8), (boxes) => {
          const moves = distributeBoxes(boxes, axis);
          expect(moves).toHaveLength(boxes.length);
          for (const move of moves) {
            expect(move.width).toBeUndefined();
            expect(move.height).toBeUndefined();
          }
          const sorted = [...boxes].sort((a, b) =>
            axis === "horizontal" ? a.left - b.left : a.top - b.top,
          );
          const moveById = new Map(moves.map((m) => [m.id, m]));
          const first = sorted[0]!;
          const last = sorted.at(-1)!;
          if (axis === "horizontal") {
            expect(moveById.get(first.id)!.left).toBeCloseTo(first.left, 6);
            expect(moveById.get(last.id)!.left! + last.width).toBeCloseTo(
              last.left + last.width,
              6,
            );
          } else {
            expect(moveById.get(first.id)!.top).toBeCloseTo(first.top, 6);
            expect(moveById.get(last.id)!.top! + last.height).toBeCloseTo(
              last.top + last.height,
              6,
            );
          }
        }),
        { seed: SEED, numRuns: NUM_RUNS },
      );
    },
  );

  it("fewer than three boxes always answers no moves at all", () => {
    fc.assert(
      fc.property(
        boxesArb(0, 2),
        fc.constantFrom("horizontal", "vertical"),
        (boxes, axis) => {
          expect(distributeBoxes(boxes, axis)).toEqual([]);
        },
      ),
      { seed: SEED, numRuns: 50 },
    );
  });
});

describe("matchSize: every target copies the reference's own size, the reference is never moved", () => {
  it("targets get the reference's width and height, non-negative, reference untouched", () => {
    fc.assert(
      fc.property(boxesArb(2, 8), (boxes) => {
        const moves = matchSize(boxes);
        const source = boxes[0]!;
        expect(moves).toHaveLength(boxes.length - 1);
        expect(moves.some((m) => m.id === source.id)).toBe(false);
        for (const move of moves) {
          expect(move.width).toBe(source.width);
          expect(move.height).toBe(source.height);
          expect(move.width!).toBeGreaterThanOrEqual(0);
          expect(move.height!).toBeGreaterThanOrEqual(0);
          // matchSize only ever resizes: no position field.
          expect(move.left).toBeUndefined();
          expect(move.top).toBeUndefined();
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("fewer than two boxes always answers no moves at all", () => {
    fc.assert(
      fc.property(boxesArb(0, 1), (boxes) => {
        expect(matchSize(boxes)).toEqual([]);
      }),
      { seed: SEED, numRuns: 50 },
    );
  });
});

describe("swapBoxes: positions trade, sizes never appear in the moves", () => {
  it("each box takes the other's left/top exactly, width and height are absent", () => {
    fc.assert(
      fc.property(boxesArb(2, 2), (boxes) => {
        const moves = swapBoxes(boxes);
        const [a, b] = boxes as [ObjectBox, ObjectBox];
        expect(moves).toEqual([
          { id: a.id, left: b.left, top: b.top },
          { id: b.id, left: a.left, top: a.top },
        ]);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("anything but exactly two boxes always answers no moves at all", () => {
    fc.assert(
      fc.property(fc.oneof(boxesArb(0, 1), boxesArb(3, 6)), (boxes) => {
        expect(swapBoxes(boxes)).toEqual([]);
      }),
      { seed: SEED, numRuns: 100 },
    );
  });
});

describe("none of the four ever produces a move with a negative size", () => {
  it("across align, distribute, matchSize and swap on the same random selections", () => {
    fc.assert(
      fc.property(boxesArb(0, 8), (boxes) => {
        const all = [
          ...alignBoxes(boxes, "center"),
          ...distributeBoxes(boxes, "horizontal"),
          ...matchSize(boxes),
          ...swapBoxes(boxes),
        ];
        for (const move of all) {
          if (move.width !== undefined)
            expect(move.width).toBeGreaterThanOrEqual(-EPS);
          if (move.height !== undefined)
            expect(move.height).toBeGreaterThanOrEqual(-EPS);
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});
