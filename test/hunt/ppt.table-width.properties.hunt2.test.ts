// Pass-2 property: columnWidths (src/ppt/tables.ts) always answers whole
// points, each at least 1, summing EXACTLY to Math.round(width) - the
// invariant PowerPoint for the web enforces on shapes.addTable (InvalidArgument
// otherwise, rig 27.09; test/fakeppt/size-guard.ts requireColumnWidthsMatch).
// Random source widths, hidden columns (Excel's own zero) included, and
// random fractional table widths, generated with a table always roomy enough
// for its own column count (a table wanting less than a point per column is a
// degenerate input this property is not chasing, the same spirit as
// rounding.properties.hunt2.test.ts's own documented filters).

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { columnWidths } from "../../src/ppt/tables";

const SEED = 20260927;
const RUNS = 3000;

const payloadOf = (widths: number[]) =>
  ({ widths, cols: widths.length }) as Parameters<typeof columnWidths>[0];

// One source column's own width, in points, as Excel reports it: 0 for a
// hidden column, weighted to come up often since that is the branch with its
// own separate code path (the all-hidden table below), or a modelling-sized
// positive width.
const columnWidthArb = fc.oneof(
  { weight: 1, arbitrary: fc.constant(0) },
  { weight: 4, arbitrary: fc.double({ min: 0.01, max: 2000, noNaN: true }) },
);

// Every column hidden - sourceWidths' own all-zero case, which columnWidths
// answers through its equal-weight fallback - on its own weight, so it comes
// up as often as the general mixed-visibility table.
const widthsArb = fc.oneof(
  {
    weight: 1,
    arbitrary: fc
      .integer({ min: 1, max: 20 })
      .map((cols) => new Array<number>(cols).fill(0)),
  },
  {
    weight: 3,
    arbitrary: fc.array(columnWidthArb, { minLength: 1, maxLength: 20 }),
  },
);

// A table's own width, always fractional-capable (fitInto never rounds); kept
// to at least a point per column, the one bound that is not a choice but a
// mathematical floor - every column keeps at least one point, so a table
// rounding to fewer points than it has columns cannot exist, whatever the
// apportionment. Below it this property would be chasing an impossible input,
// not a bug.
const caseArb = fc
  .tuple(widthsArb, fc.double({ min: 1, max: 3000, noNaN: true }))
  .filter(([widths, width]) => Math.round(width) >= widths.length);

describe("columnWidths sums exactly to the table's rounded width", () => {
  it("holds over random source widths, hidden columns included, and fractional table widths", () => {
    fc.assert(
      fc.property(caseArb, ([widths, width]) => {
        const result = columnWidths(payloadOf(widths), width);
        expect(result).toHaveLength(widths.length);
        for (const one of result) {
          expect(Number.isInteger(one)).toBe(true);
          expect(one).toBeGreaterThanOrEqual(1);
        }
        const sum = result.reduce((total, one) => total + one, 0);
        expect(sum).toBe(Math.round(width));
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("columnWidths on the all-hidden table (the zero-total branch)", () => {
  it("still sums exactly to the rounded width, in equal whole columns", () => {
    fc.assert(
      fc.property(
        fc
          .integer({ min: 1, max: 20 })
          .chain((cols) =>
            fc.tuple(
              fc.constant(cols),
              fc.double({ min: cols, max: 3000, noNaN: true }),
            ),
          ),
        ([cols, width]) => {
          const widths = new Array<number>(cols).fill(0);
          const result = columnWidths(payloadOf(widths), width);
          expect(result).toHaveLength(cols);
          expect(result.reduce((total, one) => total + one, 0)).toBe(
            Math.round(width),
          );
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
