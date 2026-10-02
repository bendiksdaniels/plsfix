// Hunt pass 2, target 3: layoutChart (src/chart-shapes.ts, the one entry
// point every chart-shapes*.ts file feeds) as a property across every chart
// kind, random series (1-40 points, 1-6 series), any sign, zeros, and huge
// or tiny magnitudes. Plain-word invariant, the one every existing
// chart-shapes suite already states in its own header comment: every
// primitive lies inside the box it was given, with a non-negative width and
// height (PowerPoint throws InvalidArgument on a negative side, on both an
// add and a later write). Pass 1's edge tests (ppt.chart-shapes.edges.hunt)
// hand-picked single-series scenarios per kind; this generates thousands of
// them, across all seven kinds and multi-series charts pass 1 never tried,
// and lets fast-check shrink any counterexample to its smallest form.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { layoutChart, MIN_SIZE, type Primitive } from "../../src/chart-shapes";
import type { Box } from "../../src/layout";
import type {
  ChartData,
  ChartKind,
  ChartSeries,
} from "../../src/link/chart-model";

const SEED = 20260927;
const NUM_RUNS = 400;
const EPS = 0.05;

const KINDS: ChartKind[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "waterfall",
  "pie",
  "line",
];

// Any sign, zero, and both huge and tiny finite magnitudes - what a real
// modeller's sheet can hold and Excel will still export as a number.
const extremeValue = fc.oneof(
  fc.integer({ min: -1_000_000, max: 1_000_000 }),
  fc.constant(0),
  fc.constant(-0),
  fc.double({ min: 1, max: 1e300, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: -1e300, max: -1, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: 1e-300, max: 1e-6, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: -1e-6, max: -1e-300, noNaN: true, noDefaultInfinity: true }),
);

const HEX_COLORS = [
  "#2EC4B6",
  "#B27E54",
  "#14213D",
  "#E63946",
  "#8AB17D",
  "#F4A259",
];

function seriesArb(points: number): fc.Arbitrary<ChartSeries> {
  return fc
    .tuple(
      fc.string({ minLength: 1, maxLength: 10 }),
      fc.array(extremeValue, { minLength: points, maxLength: points }),
      fc.array(fc.constantFrom(...HEX_COLORS), {
        minLength: points,
        maxLength: points,
      }),
    )
    .map(([name, values, colors]): ChartSeries => ({
      name,
      values,
      labels: values.map((v) => String(v)),
      colors,
    }));
}

// kind, category count and series count first, so every series generated
// afterwards shares exactly the same point count (layoutChart, like real
// Excel exports, assumes every series in one chart has the same length).
//
// OUT-OF-SLICE, not worked around here by accident: chartmath.ts's
// bridgeSeries (src/chartmath.ts:26-34, not owned by this slice) throws
// below 3 points ("A bridge needs an opening total, a delta and a closing
// total."), but link/chart-model.ts's CHART_MIN_POINTS (line 40, also not
// owned) admits a "waterfall" at the blanket minimum of 2 like every other
// kind - a real mismatch a corrupted or future payload could hit, reported
// under OUT-OF-SLICE. Generating waterfall at >= 3 here tests this file's
// OWN invariant over the inputs valueScale (chart-shapes.ts:178-179) can
// actually be called with, rather than re-failing on the other module's gap
// on every run.
const pointsFor = (kind: ChartKind): fc.Arbitrary<number> =>
  fc.integer({ min: kind === "waterfall" ? 3 : 1, max: 40 });

const chartDataArb = fc
  .tuple(
    fc.constantFrom(...KINDS),
    fc.integer({ min: 1, max: 6 }),
    fc.option(fc.string({ minLength: 1, maxLength: 20 }), { nil: null }),
  )
  .chain(([kind, seriesCount, title]) =>
    pointsFor(kind).chain((points) =>
      fc
        .array(seriesArb(points), {
          minLength: seriesCount,
          maxLength: seriesCount,
        })
        .map((series): ChartData => ({
          v: 1,
          kind,
          title,
          categories: Array.from(
            { length: points },
            (_, i) => `Cat ${String(i)}`,
          ),
          series,
          font: "Aptos Narrow",
          ink: "#282623",
          titleColor: "#14213D",
        })),
    ),
  );

// Comfortably at or above MIN_SIZE on both sides, the one precondition every
// real caller (charts.ts belowMinimum) already guarantees before a chart is
// ever laid out at all.
const boxArb = fc.record({
  left: fc.integer({ min: 0, max: 200 }),
  top: fc.integer({ min: 0, max: 200 }),
  width: fc.integer({ min: MIN_SIZE.width, max: 900 }),
  height: fc.integer({ min: MIN_SIZE.height, max: 500 }),
});

function offenders(primitives: Primitive[], box: Box): Primitive[] {
  return primitives.filter((p) => {
    const { box: pb } = p;
    return (
      !Number.isFinite(pb.left) ||
      !Number.isFinite(pb.top) ||
      !Number.isFinite(pb.width) ||
      !Number.isFinite(pb.height) ||
      pb.width < -EPS ||
      pb.height < -EPS ||
      pb.left < box.left - EPS ||
      pb.top < box.top - EPS ||
      pb.left + pb.width > box.left + box.width + EPS ||
      pb.top + pb.height > box.top + box.height + EPS
    );
  });
}

describe("layoutChart: every primitive stays inside its box, non-negative sides", () => {
  it("holds across all seven kinds, 1-40 points, 1-6 series, any sign, zero, huge and tiny values", () => {
    fc.assert(
      fc.property(chartDataArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        expect(offenders(out, box)).toEqual([]);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("draws at least one primitive whatever the kind or values (never an empty chart)", () => {
    fc.assert(
      fc.property(chartDataArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        expect(out.length).toBeGreaterThan(0);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});
