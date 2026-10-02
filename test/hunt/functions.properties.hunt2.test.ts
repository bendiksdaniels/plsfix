// Pass-2 property: PLSFIX.ROUND's results sum to PLSFIX.ROUNDSUM over the same
// range, CAGR inverts back through compounding, and every bad argument to any
// of the three custom functions answers the documented #VALUE! - never a raw
// JS exception the custom-functions runtime cannot present. Hammered with
// fast-check against the real src/functions/index.ts entry points (the
// CustomFunctions.Error stub pass-1's functions.custom.hunt.test.ts uses),
// plus the pure cagr in src/chartmath.ts for the compounding identity.
import fc from "fast-check";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cagr } from "../../src/chartmath";
import { ROUNDING_CELL_CAP, ROUNDING_MAX_DECIMALS } from "../../src/rounding";

const SEED = 20260927;
const RUNS = 2000;

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

interface Loaded {
  round: (range: unknown, index: unknown, decimals: unknown) => number;
  roundSum: (range: unknown, decimals: unknown) => number;
  cagr: (first: unknown, last: unknown, periods: unknown) => number;
}

async function loadFunctions(): Promise<Loaded> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => {
      /* not needed here */
    },
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  const module = await import("../../src/functions/index");
  return {
    round: module.smtRound as Loaded["round"],
    roundSum: module.smtRoundSum as Loaded["roundSum"],
    cagr: module.smtCagr as Loaded["cagr"],
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

function column(values: number[]): number[][] {
  return values.map((value) => [value]);
}

// Magnitude spread similar to rounding.properties.hunt2, kept safely inside a
// double's exact-integer range at the shifted decimals (rounding.ts's own
// documented reason: ROUNDING_MAX_DECIMALS's comment), so any deviation found
// is this layer's own, not the float limit rounding.properties.hunt2 already
// pins.
const magnitudeValueArb = fc
  .tuple(
    fc.boolean(),
    fc.double({ min: 1, max: 10, noNaN: true }),
    fc.integer({ min: -6, max: 6 }),
  )
  .map(([negative, mantissa, exponent]) => {
    const value = mantissa * 10 ** exponent;
    return negative ? -value : value;
  });

const decimalsArb = fc.integer({
  min: -ROUNDING_MAX_DECIMALS,
  max: ROUNDING_MAX_DECIMALS,
});

const caseArb = fc
  .tuple(
    fc.array(magnitudeValueArb, { minLength: 1, maxLength: 40 }),
    decimalsArb,
  )
  .filter(
    ([values, decimals]) =>
      values.length <= ROUNDING_CELL_CAP &&
      values.every((value) => Math.abs(value) * 10 ** decimals < 2 ** 51),
  );

describe("PLSFIX.ROUND sums to PLSFIX.ROUNDSUM over the whole range", () => {
  it("agrees for every generated range and decimals through the real custom functions", async () => {
    const { round, roundSum } = await loadFunctions();
    await fc.assert(
      fc.asyncProperty(caseArb, async ([values, decimals]) => {
        const range = column(values);
        const total = roundSum(range, decimals);
        let sum = 0;
        for (let index = 1; index <= values.length; index += 1) {
          sum += round(range, index, decimals);
        }
        const unit = 10 ** -decimals;
        expect(Math.abs(sum - total)).toBeLessThan(unit / 2);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("PLSFIX.CAGR inverts back through compounding", () => {
  it("first * (1 + rate)^periods lands back on last, within tolerance", () => {
    fc.assert(
      fc.property(
        fc.double({ min: 1e-6, max: 1e9, noNaN: true }),
        fc.double({ min: 1e-6, max: 1e9, noNaN: true }),
        fc.double({ min: 1, max: 1000, noNaN: true }),
        (first, last, periods) => {
          fc.pre(first > 0 && last > 0 && periods >= 1);
          const rate = cagr(first, last, periods);
          const rebuilt = first * (1 + rate) ** periods;
          // A relative tolerance: compounding a rate raised to up to 1000
          // amplifies whatever float noise cagr's own log-space rescue (for
          // a last/first ratio that overflows a double) already carries.
          const relativeError = Math.abs(rebuilt - last) / Math.max(last, 1);
          expect(relativeError).toBeLessThan(1e-6);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("still rescues a rate in log space once the ratio itself overflows, and it still inverts", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: -300, max: -1 }),
        fc.integer({ min: 1, max: 300 }),
        fc.double({ min: 50, max: 500, noNaN: true }),
        (firstExp, lastExp, periods) => {
          const first = 10 ** firstExp;
          const last = 10 ** lastExp;
          fc.pre(!Number.isFinite(last / first)); // exercise the overflow path
          const rate = cagr(first, last, periods);
          expect(Number.isFinite(rate)).toBe(true);
          // log-space check: log(first) + periods*log(1+rate) ~= log(last).
          const rebuiltLog = Math.log(first) + periods * Math.log(1 + rate);
          const relativeError =
            Math.abs(rebuiltLog - Math.log(last)) / Math.abs(Math.log(last));
          expect(relativeError).toBeLessThan(1e-6);
        },
      ),
      { seed: SEED, numRuns: 500 },
    );
  });
});

// Every argument shape a modeller could hand a cell function, valid or not.
const badRangeArb = fc.oneof(
  fc.constant([]),
  fc.constant([[]]),
  fc.array(fc.array(fc.double({ noNaN: true }), { maxLength: 5 }), {
    maxLength: 5,
  }),
  fc
    .array(
      fc.oneof(
        fc.double({ noNaN: true }),
        fc.constant(Number.NaN),
        fc.constant(Number.POSITIVE_INFINITY),
        fc.constant(Number.NEGATIVE_INFINITY),
        fc.boolean(),
        fc.constant(null),
        fc.string(),
      ),
      { minLength: 1, maxLength: 6 },
    )
    .map((row) => [row]),
  fc
    .array(fc.double({ noNaN: true }), {
      minLength: ROUNDING_CELL_CAP + 1,
      maxLength: ROUNDING_CELL_CAP + 50,
    })
    .map(column),
);

const badNumberArb = fc.oneof(
  fc.double(),
  fc.constant(Number.NaN),
  fc.constant(Number.POSITIVE_INFINITY),
  fc.constant(Number.NEGATIVE_INFINITY),
  fc.integer(),
  fc.constant(0),
  fc.constant(-1),
);

async function callNeverThrowsRaw(
  run: () => unknown,
): Promise<{ ok: true; value: unknown } | { ok: false; error: unknown }> {
  try {
    return { ok: true, value: run() };
  } catch (error) {
    return { ok: false, error };
  }
}

function isBadDecimals(decimals: unknown): boolean {
  return (
    typeof decimals !== "number" ||
    !Number.isInteger(decimals) ||
    Math.abs(decimals) > ROUNDING_MAX_DECIMALS
  );
}

function totalCells(range: unknown[][]): number {
  return range.reduce((total: number, row) => total + row.length, 0);
}

function isBadRange(range: unknown[][]): boolean {
  if (totalCells(range) > ROUNDING_CELL_CAP) return true;
  return range.some((row) =>
    row.some((cell) => typeof cell !== "number" || !Number.isFinite(cell)),
  );
}

describe("a bad argument always answers the documented #VALUE!, never a raw exception", () => {
  it("PLSFIX.ROUND and PLSFIX.ROUNDSUM refuse cleanly over generated bad ranges and decimals", async () => {
    const { round, roundSum } = await loadFunctions();
    await fc.assert(
      fc.asyncProperty(
        badRangeArb,
        fc.oneof(decimalsArb, badNumberArb),
        fc.integer({ min: -5, max: 45 }),
        async (range, decimals, index) => {
          const rangeBad = isBadRange(range as unknown[][]);
          const decimalsBad = isBadDecimals(decimals);

          const outcome = await callNeverThrowsRaw(() =>
            roundSum(range as number[][], decimals),
          );
          // A bad range or bad decimals is documented to refuse ROUNDSUM
          // unconditionally (src/functions/index.ts rangeValues,
          // requireDecimals) - success here would be silently wrong, not
          // just an uncaught exception.
          if (rangeBad || decimalsBad) expect(outcome.ok).toBe(false);
          if (!outcome.ok) {
            expect(outcome.error).toBeInstanceOf(FakeFunctionError);
            expect((outcome.error as FakeFunctionError).code).toBe("#VALUE!");
          }

          const roundOutcome = await callNeverThrowsRaw(() =>
            round(range as number[][], index, decimals),
          );
          const indexBad =
            !Number.isInteger(index) ||
            index < 1 ||
            (!rangeBad && index > totalCells(range as unknown[][]));
          if (rangeBad || decimalsBad || indexBad) {
            expect(roundOutcome.ok).toBe(false);
          }
          if (!roundOutcome.ok) {
            expect(roundOutcome.error).toBeInstanceOf(FakeFunctionError);
            expect((roundOutcome.error as FakeFunctionError).code).toBe(
              "#VALUE!",
            );
          }
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("PLSFIX.CAGR refuses cleanly over any generated first/last/periods", async () => {
    const { cagr: smtCagr } = await loadFunctions();
    await fc.assert(
      fc.asyncProperty(
        badNumberArb,
        badNumberArb,
        badNumberArb,
        async (first, last, periods) => {
          const outcome = await callNeverThrowsRaw(() =>
            smtCagr(first, last, periods),
          );
          // Each argument is individually documented as needing a positive,
          // finite number (src/functions/index.ts requirePositive): whenever
          // any one of the three is not, a success (any success, silently
          // returning something finite) is as wrong as a raw JS exception -
          // both hide a bad argument instead of naming it. This is the check
          // that actually distinguishes "properly refused" from "quietly
          // produced a number anyway" (confirmed by briefly removing
          // requirePositive(periods, ...) by hand: without it, the weaker
          // "finite-or-refused" version of this assertion still passed).
          const isBad = (value: number) =>
            typeof value !== "number" || !Number.isFinite(value) || value <= 0;
          if (isBad(first) || isBad(last) || isBad(periods)) {
            expect(outcome.ok).toBe(false);
          }
          if (!outcome.ok) {
            expect(outcome.error).toBeInstanceOf(FakeFunctionError);
            expect((outcome.error as FakeFunctionError).code).toBe("#VALUE!");
          } else {
            // A number CustomFunctions can actually print, never Infinity.
            expect(Number.isFinite(outcome.value as number)).toBe(true);
          }
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
