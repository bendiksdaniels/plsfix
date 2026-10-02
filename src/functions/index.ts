// The Excel-facing custom functions, =PLSFIX.ROUND, =PLSFIX.ROUNDSUM and
// =PLSFIX.CAGR, and the only module Office loads into the custom-functions
// runtime (built to dist as a standalone functions.js). It imports the pure
// allocator and the pure chart maths and nothing else: no pane code, no DOM,
// no Office.js object model.
//
// Every cell of a group passes the whole range, so Excel recalculates all of
// them whenever any value in it changes, and each one recomputes the same
// allocation and reads out its own slot - shared behaviour without shared state
// (docs/research/custom-functions.md, section 4).

import { cagr } from "../chartmath";
import {
  allocateRounded,
  ROUNDING_CELL_CAP,
  ROUNDING_MAX_DECIMALS,
  roundedTotal,
} from "../rounding";

function valueError(message: string): CustomFunctions.Error {
  return new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidValue,
    message,
  );
}

// Excel marshals a range argument as a number[][] grid (empty cells arrive as
// 0). The size is checked before the grid is walked, so a whole-column
// reference is refused instead of flattened first.
function rangeValues(range: number[][]): number[] {
  const cells = range.reduce((total, row) => total + row.length, 0);
  if (cells > ROUNDING_CELL_CAP) {
    throw valueError(
      `pls,fix rounding groups up to ${ROUNDING_CELL_CAP} cells; this range holds ${cells}.`,
    );
  }

  const values: number[] = [];
  for (const row of range) {
    for (const cell of row) {
      if (typeof cell !== "number" || !Number.isFinite(cell)) {
        throw valueError(
          "pls,fix rounding needs numbers; this range holds text.",
        );
      }
      values.push(cell);
    }
  }
  return values;
}

function requireDecimals(decimals: number): number {
  if (
    !Number.isInteger(decimals) ||
    Math.abs(decimals) > ROUNDING_MAX_DECIMALS
  ) {
    throw valueError(
      `pls,fix rounding takes a whole number of decimals between -${ROUNDING_MAX_DECIMALS} and ${ROUNDING_MAX_DECIMALS}.`,
    );
  }
  return decimals;
}

// The range's total, rounded so it equals the sum of PLSFIX.ROUND over the range.
export function smtRoundSum(range: number[][], decimals: number): number {
  return roundedTotal(rangeValues(range), requireDecimals(decimals));
}

// One position's share of the consistently rounded range, counting from 1.
export function smtRound(
  range: number[][],
  index: number,
  decimals: number,
): number {
  const values = rangeValues(range);
  if (!Number.isInteger(index) || index < 1 || index > values.length) {
    throw valueError(
      `PLSFIX.ROUND: position must be a whole number between 1 and ${values.length}.`,
    );
  }
  return allocateRounded(values, requireDecimals(decimals))[index - 1]!;
}

// A growth rate needs a positive start, a positive end and time to run in;
// Excel can hand a scalar text or an infinity too. All of them are the bad
// argument the rounding functions report, so the cell shows one #VALUE!.
function requirePositive(value: number, what: string): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw valueError(`PLSFIX.CAGR: ${what} must be a positive number.`);
  }
}

// A throw is not always an Error; String() on one that is would print its
// class name in front of the sentence, so the message is read off it first.
function reason(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// The compound annual growth rate: what one period's growth would have to be
// for `first` to reach `last` over `periods` of them.
export function smtCagr(first: number, last: number, periods: number): number {
  requirePositive(first, "the start value");
  requirePositive(last, "the end value");
  requirePositive(periods, "the number of periods");
  let rate: number;
  try {
    rate = cagr(first, last, periods);
  } catch (error) {
    // The pure maths refuses what the pane's own CAGR refuses - a period
    // shorter than one - and its sentence travels in the same error kind.
    throw valueError(`PLSFIX.CAGR: ${reason(error)}`);
  }
  // last / first overflows to Infinity once the two are far enough apart, and
  // Infinity to any power stays Infinity: the cell would hold a number Excel
  // cannot print. A refusal is the only honest answer.
  if (!Number.isFinite(rate)) {
    throw valueError(
      "PLSFIX.CAGR: these values are too far apart to give a rate.",
    );
  }
  return rate;
}

// The UX gates (ux:check, ux:sweep) load this bundle in a headless browser
// with no Office host at all, where the global CustomFunctions runtime below
// was never defined; referencing it directly would throw before the gate got
// to check anything. typeof is the one safe way to ask for a global that may
// not exist.
function associate(
  id: string,
  implementation: typeof smtRound | typeof smtRoundSum | typeof smtCagr,
): void {
  if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate(id, implementation);
  }
}

// The ids match src/functions/metadata.ts; the manifest's <Namespace> makes
// them PLSFIX.ROUND, PLSFIX.ROUNDSUM and PLSFIX.CAGR in the grid.
associate("ROUND", smtRound);
associate("ROUNDSUM", smtRoundSum);
associate("CAGR", smtCagr);
