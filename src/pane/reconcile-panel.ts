// Small pane adapter for the pure reconciliation solver. The selected answer
// is the primary output; this summary makes the result understandable without
// forcing the user to add the cells again by hand.

import { reconcileSelection } from "../excel";
import { reconcileSummary } from "../reconcile";
import { getActiveSettings } from "../settings";
import { getElement } from "../ui/dom";

// The standing line of taskpane.html's own result paragraph: what shows before
// the first search, and again once an answer has stopped being one.
const HINT =
  "Select up to 34 numeric cells. Matching cells become the selection.";

function numberFrom(id: string, label: string): number {
  const raw = getElement<HTMLInputElement>(id).value.trim();
  // Number("") is 0, which would quietly search for a zero target.
  const value = raw === "" ? Number.NaN : Number(raw);
  if (!Number.isFinite(value)) throw new Error(`${label} must be a number.`);
  return value;
}

export async function runReconciliation(): Promise<string> {
  const target = numberFrom("reconcile-target", "Target");
  const tolerance = numberFrom("reconcile-tolerance", "Tolerance");
  if (tolerance < 0) throw new Error("Tolerance must be zero or greater.");

  // The answer goes away before the search runs: a refused one leaves the
  // sheet selected on the cells the old line names, which is not the answer to
  // what was just asked.
  const line = getElement("reconcile-result");
  line.textContent = HINT;

  const result = await reconcileSelection(target, tolerance);
  line.textContent = reconcileSummary(result, getActiveSettings().language);
  return `Found ${String(result.count)} matching cells`;
}
