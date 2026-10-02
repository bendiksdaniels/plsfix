// free-space.oracle.test.ts: largestFreeBox (src/free-space.ts) against the brute force it replaced
// on 02.10.2026 (every pair of x edges times every pair of y edges, O(n^5): 12 s for 300 shapes on
// Node 22). The same rectangle, tie-break included, on random slides, the 300-shape grid and boxes
// carrying NaN or infinite edges. Invariant: the oracle below is the v2.9.23 code, unchanged.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { largestFreeBox } from "../src/free-space";
import { clear, hasArea, type Box, type Canvas } from "../src/layout";

const CANVAS: Canvas = { width: 960, height: 540 };
const MARGIN = 36;
const GAP = 12;

function uniqueSorted(values: number[]): number[] {
  return [...new Set(values)].sort((a, b) => a - b);
}

function bruteForce(
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
): Box | null {
  const [left, top] = [margin, margin];
  const [right, bottom] = [canvas.width - margin, canvas.height - margin];
  const edges = (lo: number, hi: number, pick: (b: Box) => number[]) =>
    uniqueSorted([lo, hi, ...occupied.flatMap(pick)]).filter(
      (v) => v >= lo && v <= hi,
    );
  const xs = edges(left, right, (b) => [
    b.left,
    b.left + b.width,
    b.left - gap,
    b.left + b.width + gap,
  ]);
  const ys = edges(top, bottom, (b) => [
    b.top,
    b.top + b.height,
    b.top - gap,
    b.top + b.height + gap,
  ]);
  let best: Box | null = null;
  let bestArea = 0;
  for (let i = 0; i < xs.length; i += 1)
    for (let j = i + 1; j < xs.length; j += 1)
      for (let k = 0; k < ys.length; k += 1)
        for (let l = k + 1; l < ys.length; l += 1) {
          const c = {
            left: xs[i]!,
            top: ys[k]!,
            width: xs[j]! - xs[i]!,
            height: ys[l]! - ys[k]!,
          };
          if (!hasArea(c) || !clear(c, occupied, gap)) continue;
          if (c.width * c.height > bestArea)
            [best, bestArea] = [c, c.width * c.height];
        }
  return best;
}

const boxArb = fc.record({
  left: fc.integer({ min: -100, max: 1000 }),
  top: fc.integer({ min: -100, max: 600 }),
  width: fc.integer({ min: -20, max: 500 }),
  height: fc.integer({ min: -20, max: 300 }),
});

describe("largestFreeBox against the brute force it replaced", () => {
  it("finds the same rectangle on random slides", () => {
    fc.assert(
      fc.property(
        fc.array(boxArb, { maxLength: 12 }),
        fc.integer({ min: 0, max: 24 }),
        (boxes, gap) => {
          expect(largestFreeBox(boxes, CANVAS, MARGIN, gap)).toEqual(
            bruteForce(boxes, CANVAS, MARGIN, gap),
          );
        },
      ),
      { seed: 20261002, numRuns: 400 },
    );
  });

  it("finds the same rectangle on a slide tiled with 300 shapes, and fast", () => {
    const tiles = Array.from({ length: 300 }, (_, i) => ({
      left: (i % 20) * 48,
      top: Math.floor(i / 20) * 36,
      width: 46,
      height: 34,
    }));
    const holed = tiles.filter((_, i) => i % 37 !== 5);
    for (const boxes of [tiles, holed]) {
      const started = performance.now();
      const found = largestFreeBox(boxes, CANVAS, MARGIN, GAP);
      expect(performance.now() - started).toBeLessThan(2000);
      expect(found).toEqual(bruteForce(boxes, CANVAS, MARGIN, GAP));
    }
  });

  it("keeps the first of two equal holes, as the brute force did", () => {
    // A full-width band at 250 (40 tall, gap 12) leaves 36-238 and 302-504: two
    // holes of 202 pt, so only the tie-break decides which one is returned.
    const band = [{ left: 0, top: 250, width: 960, height: 40 }];
    const found = largestFreeBox(band, CANVAS, MARGIN, GAP);
    expect(found).toEqual(bruteForce(band, CANVAS, MARGIN, GAP));
    expect(found?.top).toBe(36);
  });

  it("lets a box with NaN or infinite edges block exactly what the brute force says", () => {
    const odd: Box[][] = [
      [{ left: 100, top: Number.NaN, width: 50, height: 50 }],
      [{ left: 100, top: 100, width: 50, height: Number.NaN }],
      [{ left: Number.NaN, top: 100, width: 50, height: 50 }],
      [{ left: 100, top: Number.NEGATIVE_INFINITY, width: 50, height: 50 }],
      [{ left: 100, top: 200, width: 50, height: Number.POSITIVE_INFINITY }],
      [
        { left: 100, top: Number.POSITIVE_INFINITY, width: 50, height: 10 },
        { left: 300, top: 80, width: 40, height: 40 },
      ],
    ];
    for (const boxes of odd)
      expect(largestFreeBox(boxes, CANVAS, MARGIN, GAP)).toEqual(
        bruteForce(boxes, CANVAS, MARGIN, GAP),
      );
  });
});
