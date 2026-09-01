// Meet-in-the-middle subset solver for variance reconciliation, plus the
// house-style summary line the pane shows for a solved result. Pure and
// bounded: 34 values produce at most 131,072 candidates per half, small enough
// for an Office pane without freezing Excel.

import { formatAmount, type Language } from "./numbers";

export const RECONCILE_MAX_VALUES = 34;

interface SumChoice {
  sum: number;
  mask: number;
  count: number;
}

export interface Reconciliation {
  indices: number[];
  sum: number;
  difference: number;
}

function choices(values: number[]): SumChoice[] {
  const out: SumChoice[] = [{ sum: 0, mask: 0, count: 0 }];
  values.forEach((value, bit) => {
    const length = out.length;
    for (let index = 0; index < length; index += 1) {
      const prior = out[index]!;
      out.push({
        sum: prior.sum + value,
        mask: prior.mask | (1 << bit),
        count: prior.count + 1,
      });
    }
  });
  return out;
}

function lowerBound(sorted: SumChoice[], wanted: number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (sorted[middle]!.sum < wanted) low = middle + 1;
    else high = middle;
  }
  return low;
}

// The search looks at the three sums nearest the wanted one, so equal sums
// must collapse to their cheapest combination first: otherwise a run of equal
// sums hides the single cell behind two that add up to the same amount.
function fewestPerSum(sorted: SumChoice[]): SumChoice[] {
  const out: SumChoice[] = [];
  for (const choice of sorted) {
    const last = out[out.length - 1];
    if (last && last.sum === choice.sum) {
      if (choice.count < last.count) out[out.length - 1] = choice;
    } else {
      out.push(choice);
    }
  }
  return out;
}

function indices(mask: number, offset: number, count: number): number[] {
  const out: number[] = [];
  for (let bit = 0; bit < count; bit += 1) {
    if ((mask & (1 << bit)) !== 0) out.push(offset + bit);
  }
  return out;
}

export function solveReconciliation(
  values: number[],
  target: number,
  tolerance: number,
): Reconciliation | null {
  if (!values.every(Number.isFinite))
    throw new Error("Every value must be finite.");
  if (!Number.isFinite(target)) throw new Error("The target must be a number.");
  if (!Number.isFinite(tolerance) || tolerance < 0) {
    throw new Error("Tolerance must be zero or greater.");
  }
  if (values.length === 0) return null;
  if (values.length > RECONCILE_MAX_VALUES) {
    throw new Error(
      `Reconciliation supports up to ${String(RECONCILE_MAX_VALUES)} numeric cells.`,
    );
  }

  const cut = Math.floor(values.length / 2);
  const left = choices(values.slice(0, cut));
  const right = fewestPerSum(
    choices(values.slice(cut)).sort((a, b) => a.sum - b.sum),
  );
  let best: {
    left: SumChoice;
    right: SumChoice;
    error: number;
    count: number;
  } | null = null;

  for (const one of left) {
    const at = lowerBound(right, target - one.sum);
    for (const index of [at - 1, at, at + 1]) {
      const two = right[index];
      if (!two || (one.count === 0 && two.count === 0)) continue;
      const error = Math.abs(one.sum + two.sum - target);
      const count = one.count + two.count;
      if (
        best === null ||
        error < best.error ||
        (error === best.error && count < best.count)
      ) {
        best = { left: one, right: two, error, count };
      }
    }
  }

  if (best === null || best.error > tolerance) return null;
  const picked = [
    ...indices(best.left.mask, 0, cut),
    ...indices(best.right.mask, cut, values.length - cut),
  ];
  const sum = picked.reduce((total, index) => total + values[index]!, 0);
  return { indices: picked, sum, difference: sum - target };
}

// A difference this small is floating-point noise from the subset sum, not a
// real variance, so the summary line shows a clean zero instead.
const ZERO_DIFFERENCE = 1e-10;
// Amounts reconcile to the cent, so the line shows cents; formatAmount's
// default of one decimal would round 12.34 to 12.3 and hide a real variance.
const SUMMARY_DECIMALS = 2;

export interface ReconciliationTotals {
  count: number;
  sum: number;
  difference: number;
}

/** The pane's result line, house number style: "N cells selected · Sum X · Variance Y". */
export function reconcileSummary(
  result: ReconciliationTotals,
  language: Language,
): string {
  const difference =
    Math.abs(result.difference) < ZERO_DIFFERENCE ? 0 : result.difference;
  const sum = formatAmount(result.sum, language, SUMMARY_DECIMALS);
  const variance = formatAmount(difference, language, SUMMARY_DECIMALS);
  return `${String(result.count)} cells selected · Sum ${sum} · Variance ${variance}`;
}
