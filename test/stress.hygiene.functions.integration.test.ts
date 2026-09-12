// Stress pass, slice P2: the three custom functions with the wrong thing in
// every argument. Office marshals a matrix parameter as number[][] and a
// scalar as whatever the cell held, so text, blanks, booleans, errors and
// numbers at the edge of a double all arrive here. Every one of them must
// leave the cell showing #VALUE! with a pls,fix sentence behind it - never a
// raw TypeError, never a non-finite number Excel has to render itself.

import { afterEach, describe, expect, it, vi } from "vitest";
import { ROUNDING_CELL_CAP, ROUNDING_MAX_DECIMALS } from "../src/rounding";

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

// A fresh module graph per test: association happens at import time, so the
// runtime stub has to stand before the module is loaded.
async function loadFunctions(): Promise<Loaded> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => undefined,
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  const module = await import("../src/functions/index");
  return {
    round: module.smtRound,
    roundSum: module.smtRoundSum,
    cagr: module.smtCagr,
  };
}

// What the cell ends up showing: the error code plus the sentence behind it,
// or the number the function answered.
function shown(run: () => number): string {
  try {
    return `= ${String(run())}`;
  } catch (error) {
    if (error instanceof FakeFunctionError) {
      return `${error.code} ${String(error.message)}`;
    }
    return `RAW ${String(error)}`;
  }
}

// Excel hands a scalar parameter whatever the cell held; the signatures are
// typed number, so the wrong kinds come in past the type system.
function loose(value: unknown): number {
  return value as number;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("PLSFIX.CAGR with the wrong thing in an argument", () => {
  it("refuses text, a blank, a boolean, an error and a range", async () => {
    const { cagr } = await loadFunctions();
    const wrong = ["abc", null, undefined, true, Number.NaN, [[100]]];

    for (const value of wrong) {
      expect(shown(() => cagr(loose(value), 200, 5))).toBe(
        "#VALUE! PLSFIX.CAGR: the start value must be a positive number.",
      );
      expect(shown(() => cagr(100, loose(value), 5))).toBe(
        "#VALUE! PLSFIX.CAGR: the end value must be a positive number.",
      );
      expect(shown(() => cagr(100, 200, loose(value)))).toBe(
        "#VALUE! PLSFIX.CAGR: the number of periods must be a positive number.",
      );
    }
  });

  it("refuses zero and negative values, and an infinity", async () => {
    const { cagr } = await loadFunctions();
    for (const value of [0, -1, Number.NEGATIVE_INFINITY]) {
      expect(shown(() => cagr(value, 200, 5))).toContain("#VALUE!");
      expect(shown(() => cagr(100, value, 5))).toContain("#VALUE!");
      expect(shown(() => cagr(100, 200, value))).toContain("#VALUE!");
    }
    expect(shown(() => cagr(100, Number.POSITIVE_INFINITY, 5))).toBe(
      "#VALUE! PLSFIX.CAGR: the end value must be a positive number.",
    );
  });

  it("refuses a period shorter than one with the pane's own sentence", async () => {
    const { cagr } = await loadFunctions();
    expect(shown(() => cagr(100, 200, 0.5))).toBe(
      "#VALUE! PLSFIX.CAGR: A CAGR needs at least one period.",
    );
    expect(shown(() => cagr(100, 200, 0.999))).toBe(
      "#VALUE! PLSFIX.CAGR: A CAGR needs at least one period.",
    );
  });

  // A start small enough and an end large enough overflow the ratio to
  // Infinity, and Infinity ** anything is Infinity: the cell then holds a
  // number Excel cannot print. A refusal the modeller can read is the only
  // honest answer.
  it("refuses a ratio that overflows a double instead of answering Infinity", async () => {
    const { cagr } = await loadFunctions();
    expect(shown(() => cagr(5e-324, 1e308, 1))).toBe(
      "#VALUE! PLSFIX.CAGR: these values are too far apart to give a rate.",
    );
    expect(shown(() => cagr(1e-300, 1e300, 1))).toBe(
      "#VALUE! PLSFIX.CAGR: these values are too far apart to give a rate.",
    );
  });

  it("still answers the ordinary cases, and the flat one", async () => {
    const { cagr } = await loadFunctions();
    expect(cagr(100, 200, 5)).toBeCloseTo(0.148698, 6);
    expect(cagr(100, 100, 3)).toBe(0);
    expect(cagr(200, 100, 1)).toBe(-0.5);
    // A period count past every sensible model still has an answer: no growth
    // per period at all.
    expect(cagr(100, 200, 1e308)).toBe(0);
  });
});

describe("PLSFIX.ROUNDSUM and PLSFIX.ROUND over a range a modeller really has", () => {
  it("names the reason for text, a blank, a boolean and a NaN", async () => {
    const { roundSum } = await loadFunctions();
    for (const cell of ["x", null, true, Number.NaN]) {
      expect(shown(() => roundSum([[1, loose(cell)]], 0))).toBe(
        "#VALUE! pls,fix rounding needs numbers; this range holds text.",
      );
    }
  });

  it("refuses one cell past the group cap and takes the cap itself", async () => {
    const { roundSum } = await loadFunctions();
    const row = (count: number): number[][] => [
      Array.from({ length: count }, () => 1),
    ];

    expect(roundSum(row(ROUNDING_CELL_CAP - 1), 0)).toBe(ROUNDING_CELL_CAP - 1);
    expect(roundSum(row(ROUNDING_CELL_CAP), 0)).toBe(ROUNDING_CELL_CAP);
    expect(shown(() => roundSum(row(ROUNDING_CELL_CAP + 1), 0))).toBe(
      `#VALUE! pls,fix rounding groups up to ${String(ROUNDING_CELL_CAP)} cells; this range holds ${String(ROUNDING_CELL_CAP + 1)}.`,
    );
  });

  it("refuses a decimals argument at cap + 1, fractional or text", async () => {
    const { roundSum } = await loadFunctions();
    const sentence = `#VALUE! pls,fix rounding takes a whole number of decimals between -${String(ROUNDING_MAX_DECIMALS)} and ${String(ROUNDING_MAX_DECIMALS)}.`;

    expect(roundSum([[1.5, 1.5]], ROUNDING_MAX_DECIMALS)).toBe(3);
    expect(roundSum([[1.5, 1.5]], -ROUNDING_MAX_DECIMALS)).toBe(0);
    for (const decimals of [
      ROUNDING_MAX_DECIMALS + 1,
      -(ROUNDING_MAX_DECIMALS + 1),
      0.5,
      loose("x"),
      loose(null),
    ]) {
      expect(shown(() => roundSum([[1.5, 1.5]], decimals))).toBe(sentence);
    }
  });

  it("holds negatives, zeros, 15-digit numbers and a lone cell", async () => {
    const { round, roundSum } = await loadFunctions();
    expect(roundSum([[-1.5, -1.5]], 0)).toBe(-3);
    expect(roundSum([[0, 0]], 0)).toBe(0);
    expect(roundSum([[123456789012345, 0.5]], 0)).toBe(123456789012346);
    expect(roundSum([[7.4]], 0)).toBe(7);
    expect(round([[7.4]], 1, 0)).toBe(7);
  });

  it("names a position outside the range instead of guessing one", async () => {
    const { round } = await loadFunctions();
    for (const index of [0, -1, 3, 1.5, loose("a"), loose(null)]) {
      expect(shown(() => round([[1, 2]], index, 0))).toBe(
        "#VALUE! PLSFIX.ROUND: position must be a whole number between 1 and 2.",
      );
    }
  });

  // Office marshals every matrix parameter as number[][], a scalar reference
  // included, so the empty grid is the one shape a bare reference could take.
  it("answers an empty group without a raw error", async () => {
    const { round, roundSum } = await loadFunctions();
    expect(roundSum([], 0)).toBe(0);
    expect(roundSum([[]], 0)).toBe(0);
    expect(shown(() => round([[]], 1, 0))).toContain("#VALUE!");
    expect(shown(() => round([[]], 1, 0))).not.toContain("RAW");
  });

  // The whole point of the pair: every part of the group and the total agree,
  // whatever the cells hold.
  it("keeps the parts adding up to the total on an awkward group", async () => {
    const { round, roundSum } = await loadFunctions();
    const group = [[33.333], [33.333], [33.334]];
    const parts = [1, 2, 3].map((index) => round(group, index, 0));

    expect(parts.reduce((sum, part) => sum + part, 0)).toBe(
      roundSum(group, 0),
    );
  });
});
