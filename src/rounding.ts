// Consistent rounding, think-cell TCROUND style: the rounded parts of a group
// add up to the rounded total, instead of missing it by a unit. Largest
// remainder - every part is floored, and the units still owed to the total go
// to the parts that lost the most, ties by position.
//
// Pure by design and imports nothing: the custom-functions bundle
// (src/functions/) loads this module into Excel's function runtime, where none
// of the pane's DOM or Office.js code may follow it.

// Above this many cells the group stops being free to recalculate: every
// sibling SMT.ROUND cell redoes the whole allocation, so one edit costs
// O(N^2 log N). Reconciliation blocks - the case this is written for - are
// dozens of rows, not thousands (docs/research/custom-functions.md, Risks).
export const ROUNDING_CELL_CAP = 1_000;

// 10^10 is the last power of ten that scales a modelling-sized number without
// leaving the exact-integer range of a double.
export const ROUNDING_MAX_DECIMALS = 10;

// Multiplying by 10^decimals introduces binary noise a modeller never typed
// (1.005 * 100 is 100.49999999999999), which would reorder remainders. Shifting
// the exponent in the decimal string instead is exact: "1.005e2" parses to
// 100.5. Values already in exponent form ("1e+21") keep their exponent.
function shiftDecimal(value: number, places: number): number {
  const [mantissa, exponent] = String(value).split("e");
  return Number(`${mantissa}e${Number(exponent ?? 0) + places}`);
}

// Excel's own ROUND: halves go away from zero, so -0.5 rounds to -1, not to 0.
function roundHalfAway(units: number): number {
  return Math.sign(units) * Math.round(Math.abs(units));
}

function assertInputs(values: number[], decimals: number): void {
  if (
    !Number.isInteger(decimals) ||
    Math.abs(decimals) > ROUNDING_MAX_DECIMALS
  ) {
    throw new Error(
      `rounding: decimals must be a whole number between -${ROUNDING_MAX_DECIMALS} and ${ROUNDING_MAX_DECIMALS}`,
    );
  }
  if (values.some((value) => !Number.isFinite(value))) {
    throw new Error("rounding: every value must be a finite number");
  }
}

// The grand total in whole units of 10^-decimals: what the parts have to add
// up to, and the one place the target is decided for both exports.
function totalUnits(values: number[], decimals: number): number {
  const total = values.reduce((sum, value) => sum + value, 0);
  return roundHalfAway(shiftDecimal(total, decimals));
}

// The rounded total of the group: SMT.ROUNDSUM, and by construction the sum of
// what allocateRounded returns for the same values.
export function roundedTotal(values: number[], decimals: number): number {
  assertInputs(values, decimals);
  return shiftDecimal(totalUnits(values, decimals), -decimals);
}

// Indices ordered by what each value lost to the floor, largest first; equal
// remainders keep their position order, so the allocation is the same for every
// sibling cell that recomputes it.
function bumpOrder(remainders: number[]): number[] {
  return remainders
    .map((remainder, index) => ({ remainder, index }))
    .sort((a, b) => b.remainder - a.remainder || a.index - b.index)
    .map((entry) => entry.index);
}

// Each value rounded to `decimals`, with the whole set adding up to
// roundedTotal(values, decimals) exactly. Deterministic: same inputs, same
// output, in any cell that asks.
export function allocateRounded(values: number[], decimals: number): number[] {
  assertInputs(values, decimals);
  if (values.length === 0) return [];

  const units = values.map((value) => shiftDecimal(value, decimals));
  const floors = units.map((unit) => Math.floor(unit));
  const floored = floors.reduce((sum, floor) => sum + floor, 0);
  // Whole units the floors still owe the total: between 0 and one per value,
  // clamped because a float sum of very large inputs could drift past either
  // end and there is no such thing as bumping a value that is not there.
  const owed = Math.min(
    values.length,
    Math.max(0, Math.round(totalUnits(values, decimals) - floored)),
  );

  const bumped = new Set(
    bumpOrder(units.map((unit, index) => unit - floors[index]!)).slice(0, owed),
  );
  return floors.map((floor, index) =>
    shiftDecimal(floor + (bumped.has(index) ? 1 : 0), -decimals),
  );
}
