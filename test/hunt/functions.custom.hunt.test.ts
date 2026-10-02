// Pass-1 hunt: =PLSFIX.ROUND / =PLSFIX.ROUNDSUM / =PLSFIX.CAGR as Excel calls
// them. Focus: the decimals boundary through the custom-function layer (not
// just the pure allocator), the literal 10,000-cell refusal, a non-number cell
// that is not text, and CAGR's overflow rescue in log space.
import { afterEach, describe, expect, it, vi } from "vitest";
import { ROUNDING_CELL_CAP, ROUNDING_MAX_DECIMALS } from "../../src/rounding";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

interface Loaded {
  round: (range: number[][], index: number, decimals: number) => number;
  roundSum: (range: number[][], decimals: number) => number;
  cagr: (first: number, last: number, periods: number) => number;
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
    round: module.smtRound,
    roundSum: module.smtRoundSum,
    cagr: module.smtCagr,
  };
}

function column(values: number[]): number[][] {
  return values.map((value) => [value]);
}

async function rejects(run: () => unknown): Promise<FakeFunctionError> {
  try {
    run();
  } catch (error) {
    return error as FakeFunctionError;
  }
  throw new Error("expected a #VALUE! error");
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("the decimals boundary through the custom function", () => {
  it("allows exactly the max decimals in both directions", async () => {
    const { round, roundSum } = await loadFunctions();
    const range = column([1, 2]);
    expect(round(range, 1, ROUNDING_MAX_DECIMALS)).toBe(1);
    expect(round(range, 1, -ROUNDING_MAX_DECIMALS)).toBe(0);
    expect(roundSum(range, ROUNDING_MAX_DECIMALS)).toBe(3);
    expect(roundSum(range, -ROUNDING_MAX_DECIMALS)).toBe(0);
  });

  it("refuses one past the max decimals on the negative side too", async () => {
    const { round, roundSum } = await loadFunctions();
    const range = column([1, 2]);
    const error = await rejects(() =>
      round(range, 1, -ROUNDING_MAX_DECIMALS - 1),
    );
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toContain("whole number of decimals");
    expect(
      (await rejects(() => roundSum(range, -ROUNDING_MAX_DECIMALS - 1))).code,
    ).toBe("#VALUE!");
  });
});

describe("the literal ten-thousand-cell range the brief names", () => {
  it("names the actual count over the cap", async () => {
    const { round, roundSum } = await loadFunctions();
    const oversized = column(Array.from({ length: 10_000 }, () => 1));

    const error = await rejects(() => round(oversized, 1, 0));
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toContain(`up to ${ROUNDING_CELL_CAP} cells`);
    expect(error.message).toContain("holds 10000");
    expect((await rejects(() => roundSum(oversized, 0))).message).toContain(
      "holds 10000",
    );
  });
});

describe("a range cell that is not a number and not text either", () => {
  it("refuses a boolean cell instead of treating it as a number", async () => {
    const { roundSum } = await loadFunctions();
    const range = [[1], [true]] as unknown as number[][];
    const error = await rejects(() => roundSum(range, 0));
    expect(error.code).toBe("#VALUE!");
  });

  it("refuses a null cell instead of throwing a TypeError", async () => {
    const { roundSum } = await loadFunctions();
    const range = [[1], [null]] as unknown as number[][];
    const error = await rejects(() => roundSum(range, 0));
    expect(error.code).toBe("#VALUE!");
  });
});

describe("an empty range", () => {
  it("sums an empty range to zero", async () => {
    const { roundSum } = await loadFunctions();
    expect(roundSum([[]], 0)).toBe(0);
    expect(roundSum([], 0)).toBe(0);
  });

  it("refuses a position into an empty range without throwing a raw error", async () => {
    const { round } = await loadFunctions();
    const error = await rejects(() => round([[]], 1, 0));
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toContain("between 1 and 0");
  });
});

describe("CAGR: fractional periods and the overflow rescue", () => {
  it("allows a fractional period count of at least one", async () => {
    const { cagr } = await loadFunctions();
    // 2.5 periods is a valid stretch of time, just not a whole one.
    const rate = cagr(100, 200, 2.5);
    expect(rate).toBeCloseTo(200 ** (1 / 2.5) / 100 ** (1 / 2.5) - 1, 10);
  });

  it("rescues a rate in log space once the plain ratio overflows a double", async () => {
    const { cagr } = await loadFunctions();
    // last/first here is 1e600, which overflows to Infinity; the log-space
    // path the CLAUDE.md map names still answers a finite, sane rate.
    const rate = cagr(1e-300, 1e300, 100);
    expect(Number.isFinite(rate)).toBe(true);
    expect(rate).toBeGreaterThan(0);
    expect(rate).toBeCloseTo(999_999, -2);
  });

  it("still refuses cleanly once even the log-space rescue cannot print", async () => {
    const { cagr } = await loadFunctions();
    // One period is nowhere near enough time to make 1e600 of growth
    // printable, in log space or otherwise: a clean #VALUE!, not Infinity.
    const error = await rejects(() => cagr(1e-300, 1e300, 1));
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toContain("too far apart");
  });
});
