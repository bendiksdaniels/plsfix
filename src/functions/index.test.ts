// The custom functions as Excel calls them: the runtime is stubbed, the module
// is imported for its side effect (association), and the associated functions
// are called with the grids Excel would marshal. Also the golden check that the
// committed public/functions.json is what the typed spec renders.
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CUSTOM_FUNCTIONS, functionsMetadata } from "./metadata";
import { ROUNDING_CELL_CAP } from "../rounding";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

interface Loaded {
  associated: Map<string, unknown>;
  round: (range: number[][], index: number, decimals: number) => number;
  roundSum: (range: number[][], decimals: number) => number;
}

// A fresh module graph per test: association happens at import time, so the
// stub has to stand before the module is loaded.
async function loadFunctions(): Promise<Loaded> {
  const associated = new Map<string, unknown>();
  vi.stubGlobal("CustomFunctions", {
    associate: (id: string, implementation: unknown) => {
      associated.set(id, implementation);
    },
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  const module = await import("./index");
  return {
    associated,
    round: module.smtRound,
    roundSum: module.smtRoundSum,
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

describe("custom functions", () => {
  it("associates exactly the ids the metadata registers", async () => {
    const { associated, round, roundSum } = await loadFunctions();
    expect([...associated.keys()].sort()).toEqual(
      CUSTOM_FUNCTIONS.map((entry) => entry.id).sort(),
    );
    expect(associated.get("ROUND")).toBe(round);
    expect(associated.get("ROUNDSUM")).toBe(roundSum);
  });

  it("hands each position its share of the rounded total", async () => {
    const { round, roundSum } = await loadFunctions();
    const range = column([33.333, 33.333, 33.334]);

    const parts = [1, 2, 3].map((index) => round(range, index, 0));
    expect(parts).toEqual([33, 33, 34]);
    expect(parts.reduce((sum, part) => sum + part, 0)).toBe(roundSum(range, 0));
    expect(roundSum(range, 0)).toBe(100);
  });

  it("reads a row range and honours the decimals", async () => {
    const { round, roundSum } = await loadFunctions();
    const range = [[1.005, 2.005, 3.005]];
    expect([1, 2, 3].map((index) => round(range, index, 2))).toEqual([
      1.01, 2.01, 3,
    ]);
    expect(roundSum(range, 2)).toBe(6.02);
  });

  it("refuses a range past the group ceiling", async () => {
    const { round, roundSum } = await loadFunctions();
    const oversized = column(
      Array.from({ length: ROUNDING_CELL_CAP + 1 }, () => 1),
    );

    const error = await rejects(() => round(oversized, 1, 0));
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toContain(`up to ${ROUNDING_CELL_CAP} cells`);
    expect((await rejects(() => roundSum(oversized, 0))).code).toBe("#VALUE!");

    // The ceiling itself still works.
    expect(
      round(column(Array.from({ length: ROUNDING_CELL_CAP }, () => 1)), 1, 0),
    ).toBe(1);
  });

  it("refuses a position outside the range", async () => {
    const { round } = await loadFunctions();
    const range = column([1.5, 2.5]);
    expect((await rejects(() => round(range, 0, 0))).message).toContain(
      "between 1 and 2",
    );
    expect((await rejects(() => round(range, 3, 0))).message).toContain(
      "between 1 and 2",
    );
    expect((await rejects(() => round(range, 1.5, 0))).code).toBe("#VALUE!");
  });

  it("refuses text in the range and a fractional precision", async () => {
    const { round, roundSum } = await loadFunctions();
    const text = [[1], ["n/a"]] as unknown as number[][];
    expect((await rejects(() => roundSum(text, 0))).message).toContain(
      "needs numbers",
    );
    expect(
      (await rejects(() => round(column([1, 2]), 1, 0.5))).message,
    ).toContain("whole number of decimals");
  });
});

describe("functions metadata", () => {
  it("matches the committed public/functions.json", () => {
    const committed = readFileSync(
      new URL("../../public/functions.json", import.meta.url),
      "utf8",
    );
    expect(functionsMetadata()).toBe(committed);
  });

  it("declares the range as a matrix and the result as a scalar number", () => {
    for (const entry of CUSTOM_FUNCTIONS) {
      expect(entry.id).toBe(entry.name);
      expect(entry.result).toEqual({
        type: "number",
        dimensionality: "scalar",
      });
      expect(entry.parameters[0]).toMatchObject({
        name: "range",
        dimensionality: "matrix",
      });
      expect(entry.parameters.at(-1)).toMatchObject({ name: "decimals" });
    }
  });
});
