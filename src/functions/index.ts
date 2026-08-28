// The Excel-facing custom functions, =SMT.ROUND and =SMT.ROUNDSUM, and the only
// module Office loads into the custom-functions runtime (built to dist as a
// standalone functions.js). It imports the pure allocator and nothing else: no
// pane code, no DOM, no Office.js object model.
//
// Every cell of a group passes the whole range, so Excel recalculates all of
// them whenever any value in it changes, and each one recomputes the same
// allocation and reads out its own slot - shared behaviour without shared state
// (docs/research/custom-functions.md, section 4).

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
      `SMT rounding groups up to ${ROUNDING_CELL_CAP} cells; this range holds ${cells}.`,
    );
  }

  const values: number[] = [];
  for (const row of range) {
    for (const cell of row) {
      if (typeof cell !== "number" || !Number.isFinite(cell)) {
        throw valueError("SMT rounding needs numbers; this range holds text.");
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
      `SMT rounding takes a whole number of decimals between -${ROUNDING_MAX_DECIMALS} and ${ROUNDING_MAX_DECIMALS}.`,
    );
  }
  return decimals;
}

// The range's total, rounded so it equals the sum of SMT.ROUND over the range.
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
      `SMT.ROUND: position must be a whole number between 1 and ${values.length}.`,
    );
  }
  return allocateRounded(values, requireDecimals(decimals))[index - 1]!;
}

// The ids match src/functions/metadata.ts; the manifest's <Namespace> makes
// them SMT.ROUND and SMT.ROUNDSUM in the grid.
CustomFunctions.associate("ROUND", smtRound);
CustomFunctions.associate("ROUNDSUM", smtRoundSum);
