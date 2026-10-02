// Hunt pass 2, target 3: placeInFreeSpace (src/free-space.ts) had no direct
// unit test at all before this file - every existing test reaches it through
// placement.ts's Office.js adapter. Plain-word invariants: when it reports
// overlapping false, the box it hands back never actually overlaps any
// occupied box (with the real gap honoured) and stays within the canvas's
// margin on every side; it never shrinks the object below the caller's own
// minScale while still reporting a fit; every returned box is finite with a
// non-negative width and height. fast-check generates the occupied boxes,
// the object size and minScale; seed and run count are fixed for speed.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { placeInFreeSpace } from "../../src/free-space";
import { overlaps, type Box, type Canvas } from "../../src/layout";

const SEED = 20260927;
const NUM_RUNS = 500;
const EPS = 0.6; // rounding: placeInFreeSpace rounds every coordinate to 1pt

// The app's one real canvas and margin/gap (src/ppt/placement.ts SLIDE,
// SLIDE_MARGIN, SLIDE_GAP) - fixed, so the property tests the actual
// configuration every insert runs under, and randomises what actually
// varies: what else is already on the slide, and the size going in.
const CANVAS: Canvas = { width: 960, height: 540 };
const MARGIN = 36;
const GAP = 12;
const CONTENT_WIDTH = CANVAS.width - 2 * MARGIN;
const CONTENT_HEIGHT = CANVAS.height - 2 * MARGIN;

// Every real caller fits its size to the content area before ever reaching
// this scan (charts to CONTENT_WIDTH/height, tables and pictures to
// CONTENT_WIDTH, text likewise) - an object already bigger than the slide is
// a different, accepted case (the centred/overlapping fallback with no
// scaling at all), not this property's concern.
const sizeArb = fc.record({
  width: fc.integer({ min: 20, max: CONTENT_WIDTH }),
  height: fc.integer({ min: 20, max: CONTENT_HEIGHT }),
});

// Occupied boxes: real geometry can sit off-slide (a shape dragged out) and
// can be degenerate (zero area, from a group whose own box collapses) - both
// are things resolveTarget's own occupiedBoxes() can hand this function.
const occupiedBoxArb = fc.record({
  left: fc.integer({ min: -100, max: 1000 }),
  top: fc.integer({ min: -100, max: 600 }),
  width: fc.integer({ min: 0, max: 700 }),
  height: fc.integer({ min: 0, max: 500 }),
});

const occupiedArb = fc.array(occupiedBoxArb, { minLength: 0, maxLength: 6 });
const minScaleArb = fc.constantFrom(0.4, 0.5, 0.6, 1);

function isFiniteBox(box: Box): boolean {
  return (
    Number.isFinite(box.left) &&
    Number.isFinite(box.top) &&
    Number.isFinite(box.width) &&
    Number.isFinite(box.height)
  );
}

describe("placeInFreeSpace: overlapping false means no actual overlap, ever", () => {
  it("the returned box clears every occupied box by the real gap", () => {
    fc.assert(
      fc.property(
        sizeArb,
        occupiedArb,
        minScaleArb,
        (size, occupied, minScale) => {
          const placement = placeInFreeSpace(
            size,
            occupied,
            CANVAS,
            MARGIN,
            GAP,
            minScale,
          );
          expect(isFiniteBox(placement.box)).toBe(true);
          expect(placement.box.width).toBeGreaterThanOrEqual(0);
          expect(placement.box.height).toBeGreaterThanOrEqual(0);
          if (placement.overlapping) return;
          for (const other of occupied) {
            expect(overlaps(placement.box, other, GAP)).toBe(false);
          }
        },
      ),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("stays within the canvas's margin on every side", () => {
    fc.assert(
      fc.property(
        sizeArb,
        occupiedArb,
        minScaleArb,
        (size, occupied, minScale) => {
          const placement = placeInFreeSpace(
            size,
            occupied,
            CANVAS,
            MARGIN,
            GAP,
            minScale,
          );
          if (placement.overlapping) return;
          const { box } = placement;
          expect(box.left).toBeGreaterThanOrEqual(MARGIN - EPS);
          expect(box.top).toBeGreaterThanOrEqual(MARGIN - EPS);
          expect(box.left + box.width).toBeLessThanOrEqual(
            CANVAS.width - MARGIN + EPS,
          );
          expect(box.top + box.height).toBeLessThanOrEqual(
            CANVAS.height - MARGIN + EPS,
          );
        },
      ),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("never reports a fit scaled below the caller's own minScale", () => {
    fc.assert(
      fc.property(
        sizeArb,
        occupiedArb,
        minScaleArb,
        (size, occupied, minScale) => {
          const placement = placeInFreeSpace(
            size,
            occupied,
            CANVAS,
            MARGIN,
            GAP,
            minScale,
          );
          if (placement.overlapping) return;
          expect(placement.scale).toBeGreaterThanOrEqual(minScale - 1e-9);
          expect(placement.scale).toBeLessThanOrEqual(1 + 1e-9);
        },
      ),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("with nothing occupied, centres the object at full size, never overlapping", () => {
    fc.assert(
      fc.property(sizeArb, minScaleArb, (size, minScale) => {
        const placement = placeInFreeSpace(
          size,
          [],
          CANVAS,
          MARGIN,
          GAP,
          minScale,
        );
        expect(placement.overlapping).toBe(false);
        expect(placement.scale).toBe(1);
        expect(placement.box.width).toBe(size.width);
        expect(placement.box.height).toBe(size.height);
      }),
      { seed: SEED, numRuns: 100 },
    );
  });

  it("when it does overlap, the object is centred at full size (never a shrunken overlap)", () => {
    fc.assert(
      fc.property(
        sizeArb,
        occupiedArb,
        minScaleArb,
        (size, occupied, minScale) => {
          const placement = placeInFreeSpace(
            size,
            occupied,
            CANVAS,
            MARGIN,
            GAP,
            minScale,
          );
          if (!placement.overlapping) return;
          expect(placement.scale).toBe(1);
          expect(placement.box.width).toBe(size.width);
          expect(placement.box.height).toBe(size.height);
        },
      ),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});
