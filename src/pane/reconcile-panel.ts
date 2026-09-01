// Small pane adapter for the pure reconciliation solver. The selected answer
// is the primary output; this summary makes the result understandable without
// forcing the user to add the cells again by hand.

import { reconcileSelection } from "../excel";
import { getElement } from "../ui/dom";

function numberFrom(id: string, label: string): number {
  const value = Number(getElement<HTMLInputElement>(id).value);
  if (!Number.isFinite(value)) throw new Error(`${label} must be a number.`);
  return value;
}

export async function runReconciliation(): Promise<string> {
  const target = numberFrom("reconcile-target", "Target");
  const tolerance = numberFrom("reconcile-tolerance", "Tolerance");
  if (tolerance < 0) throw new Error("Tolerance must be zero or greater.");

  const result = await reconcileSelection(target, tolerance);
  const difference =
    Math.abs(result.difference) < 1e-10 ? 0 : result.difference;
  getElement("reconcile-result").textContent =
    `${String(result.count)} cells selected · Sum ${result.sum.toLocaleString()} · Variance ${difference.toLocaleString()}`;
  return `Found ${String(result.count)} matching cells`;
}
