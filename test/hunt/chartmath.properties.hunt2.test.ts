// Pass-2 property: src/chartmath.ts's four pure builders. bridgeSeries's rise
// minus fall reproduces the point, deltas sum to end minus start.
// tornadoSeries ranks by swing, ties stable. footballField keeps low <= high
// and flags a swap. cagr inverts the compounding it measures.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  bridgeSeries,
  cagr,
  footballField,
  formatCagrLabel,
  tornadoSeries,
} from "../../src/chartmath";

const SEED = 20260927;
const RUNS = 3000;

// Sign, mantissa in [1,10) and a power-of-ten exponent: the same spread
// rounding.properties.hunt2.test.ts uses, so small and large amounts are both
// well represented instead of a single fc.double() range that rarely samples
// the small end.
const magnitudeArb = fc
  .tuple(
    fc.boolean(),
    fc.double({ min: 1, max: 10, noNaN: true }),
    fc.integer({ min: -6, max: 9 }),
  )
  .map(([negative, mantissa, exponent]) => {
    const value = mantissa * 10 ** exponent;
    return negative ? -value : value;
  });
const amountArb = fc.oneof(
  { weight: 1, arbitrary: fc.constant(0) },
  { weight: 4, arbitrary: magnitudeArb },
);

describe("bridgeSeries: rise minus fall reproduces the original point", () => {
  it("holds for every index of every generated bridge, any sign pattern", () => {
    fc.assert(
      fc.property(
        fc.array(amountArb, { minLength: 3, maxLength: 40 }),
        (values) => {
          const bridge = bridgeSeries(values);
          values.forEach((value, index) => {
            const net = bridge.rise[index]! - bridge.fall[index]!;
            expect(net).toBeCloseTo(value, 9);
            // A middle point (a delta, not a total) always splits into a
            // non-negative rise or a non-negative fall, never both: the
            // magnitude-on-a-floor shape a waterfall's middle bars need. The
            // first and last points are totals and are plotted as
            // themselves, sign included, so this only holds strictly between
            // them.
            if (index > 0 && index < values.length - 1) {
              const rise = bridge.rise[index]!;
              const fall = bridge.fall[index]!;
              expect(rise).toBeGreaterThanOrEqual(0);
              expect(fall).toBeGreaterThanOrEqual(0);
              expect(rise === 0 || fall === 0).toBe(true);
            }
          });
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("bridgeSeries: a reconciling bridge's deltas sum to end minus start", () => {
  it("lands the implied level of the last delta on the stated closing total", () => {
    fc.assert(
      fc.property(
        fc.double({ min: -1e9, max: 1e9, noNaN: true }),
        fc.array(fc.double({ min: -1e6, max: 1e6, noNaN: true }), {
          minLength: 1,
          maxLength: 30,
        }),
        (opening, deltas) => {
          const closing = deltas.reduce((sum, delta) => sum + delta, opening);
          const values = [opening, ...deltas, closing];

          const bridge = bridgeSeries(values);
          const last = values.length - 1;
          const implied = bridge.base[last - 1]! + bridge.rise[last - 1]!;
          // The same reconciliation insertWaterfall performs on its own
          // output (src/excel/charts.ts): the deltas' running level after the
          // second-to-last point must reach the stated closing total, within
          // the float slack of summing up to 30 already-added deltas.
          const slack = Math.max(1e-6, Math.abs(closing) * 1e-9);
          expect(Math.abs(implied - closing)).toBeLessThanOrEqual(slack);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("tornadoSeries: ranked by swing, ties keep their original order", () => {
  const driverArb = fc.record({
    low: fc.double({ min: -1e6, max: 1e6, noNaN: true }),
    high: fc.double({ min: -1e6, max: 1e6, noNaN: true }),
  });
  const caseArb = fc.record({
    drivers: fc.array(driverArb, { minLength: 2, maxLength: 25 }),
    base: fc.option(fc.double({ min: -1e6, max: 1e6, noNaN: true }), {
      nil: null,
    }),
  });

  it("never raises a driver with a smaller swing above one with a bigger one", () => {
    fc.assert(
      fc.property(caseArb, ({ drivers, base }) => {
        // A unique, order-preserving label lets the ranked output be traced
        // back to the exact original driver a plain "Driver" label could not.
        const tagged = drivers.map((driver, index) => ({
          ...driver,
          label: `D${String(index)}`,
        }));
        const series = tornadoSeries(tagged, base);
        const originalIndex = new Map(
          tagged.map((driver, index) => [driver.label, index]),
        );
        const swingOf = (index: number): number =>
          Math.abs(tagged[index]!.high - tagged[index]!.low);

        // The sort's own comparator is one exact subtraction, no epsilon
        // (Math.abs(right.high - right.low) - Math.abs(left.high - left.low)):
        // the same two swings computed here the same way are bit-identical to
        // what the sort compared, so the check against them is exact too.
        const rankedIndices = series.labels.map((label) =>
          originalIndex.get(label)!,
        );
        for (let i = 1; i < rankedIndices.length; i += 1) {
          const previous = swingOf(rankedIndices[i - 1]!);
          const current = swingOf(rankedIndices[i]!);
          expect(current).toBeLessThanOrEqual(previous);
          // A stable sort never lets a later driver jump ahead of an earlier
          // one once their swing ties exactly.
          if (current === previous) {
            expect(rankedIndices[i]!).toBeGreaterThan(rankedIndices[i - 1]!);
          }
        }

        const middle =
          base ??
          tagged
            .flatMap((driver) => [driver.low, driver.high])
            .reduce((sum, value, _index, all) => sum + value / all.length, 0);
        rankedIndices.forEach((original, rankedAt) => {
          expect(series.low[rankedAt]).toBeCloseTo(
            tagged[original]!.low - middle,
            6,
          );
          expect(series.high[rankedAt]).toBeCloseTo(
            tagged[original]!.high - middle,
            6,
          );
        });
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("footballField: low never exceeds high, the band is high minus low", () => {
  const rowArb = fc.record({
    low: fc.double({ min: -1e6, max: 1e6, noNaN: true }),
    high: fc.double({ min: -1e6, max: 1e6, noNaN: true }),
  });

  it("flags exactly the rows whose low arrived above their high", () => {
    fc.assert(
      fc.property(fc.array(rowArb, { minLength: 2, maxLength: 25 }), (rows) => {
        const tagged = rows.map((row, index) => ({
          ...row,
          label: `R${String(index)}`,
        }));
        const field = footballField(tagged);

        field.low.forEach((low, index) => {
          const high = field.high[index]!;
          expect(low).toBeLessThanOrEqual(high);
          expect(field.range[index]).toBeCloseTo(high - low, 9);
        });

        const expectedSwaps = tagged.filter((row) => row.low > row.high).length;
        expect(field.swapped).toEqual(tagged.map((row) => row.low > row.high));
        expect(field.swaps).toBe(expectedSwaps);
        // Row order is preserved - a football field reads top-down in
        // selection order, never resorted by width or label.
        expect(field.labels).toEqual(tagged.map((row) => row.label));
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("cagr: the inverse of the compounding it measures", () => {
  it("recovers the rate a value was grown at over N periods", () => {
    fc.assert(
      fc.property(
        fc.double({ min: 1e-3, max: 1e9, noNaN: true }),
        fc.double({ min: -0.9, max: 3, noNaN: true }),
        fc.integer({ min: 1, max: 100 }),
        (first, rate, periods) => {
          const last = first * (1 + rate) ** periods;
          fc.pre(Number.isFinite(last) && last > 0);
          const result = cagr(first, last, periods);
          const slack = Math.max(1e-9, Math.abs(rate) * 1e-6);
          expect(Math.abs(result - rate)).toBeLessThanOrEqual(slack);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("falls back to log space when the ratio overflows a double, never throwing", () => {
    // first and last chosen so first/last overflows to Infinity in plain
    // division (1e-300 to 1e300 over 100 periods, the module's own comment).
    fc.assert(
      fc.property(
        fc.double({ min: 1e-300, max: 1e-250, noNaN: true }),
        fc.double({ min: 1e250, max: 1e300, noNaN: true }),
        fc.integer({ min: 1, max: 200 }),
        (first, last, periods) => {
          fc.pre(!Number.isFinite(last / first));
          expect(() => cagr(first, last, periods)).not.toThrow();
          const result = cagr(first, last, periods);
          // A genuine rate here is astronomical; the function is only
          // promised to answer, not to keep the answer finite.
          expect(Number.isNaN(result)).toBe(false);
        },
      ),
      { seed: SEED, numRuns: 300 },
    );
  });

  it("refuses a non-positive start, end or period count", () => {
    fc.assert(
      fc.property(
        fc.double({ min: -1e6, max: 1e6, noNaN: true }),
        fc.double({ min: -1e6, max: 1e6, noNaN: true }),
        fc.integer({ min: -10, max: 10 }),
        (first, last, periods) => {
          fc.pre(first <= 0 || last <= 0 || periods < 1);
          expect(() => cagr(first, last, periods)).toThrow();
        },
      ),
      { seed: SEED, numRuns: 500 },
    );
  });
});

describe("formatCagrLabel: never prints a negative zero", () => {
  it("shows +0.0% for any value that rounds to zero, never -0.0%", () => {
    fc.assert(
      fc.property(
        fc.double({ min: -0.00049, max: 0.00049, noNaN: true }),
        (value) => {
          fc.pre(Math.round(value * 1000) === 0);
          expect(formatCagrLabel(value)).toBe("CAGR +0.0%");
        },
      ),
      { seed: SEED, numRuns: 500 },
    );
  });

  it("always matches 'CAGR <sign><digits>.<digit>%', sign matching the rounded value", () => {
    fc.assert(
      fc.property(magnitudeArb, (value) => {
        fc.pre(Number.isFinite(value));
        const label = formatCagrLabel(value);
        const match = /^CAGR ([+-])(\d+\.\d)%$/.exec(label);
        expect(match).not.toBeNull();
        const percent = Math.round(value * 1000) / 10;
        const sign = match![1];
        expect(sign).toBe(percent < 0 ? "-" : "+");
        expect(Number(match![2])).toBeCloseTo(Math.abs(percent), 6);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
