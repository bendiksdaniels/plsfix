// Cross-cutting pane state every tab reaches for: the busy toggle, the error
// formatter, the selection-summary refresh, the tab switcher and the one
// action guard every button and ribbon command runs through. No module under
// src/pane/ may import ../main, so anything more than one tab needs lives here.

import {
  copySourceLabel,
  inspectSelection,
  lastUndoSkipped,
  undoTarget,
} from "../excel";
import { getElement } from "../ui/dom";
import { makeGuard } from "../ui/guard";
import { describeError } from "../ui/report";
import { installTabs } from "../ui/tabs";
import { createToast } from "../ui/toast";
import { formatVersion } from "../ui/version";

// Every describeError call and the footer read this one formatted constant.
export const APP_VERSION = formatVersion(__APP_VERSION__);

export const actionButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>("[data-action]"),
);
export const toast = createToast(getElement("toast"));
export const tabs = installTabs(getElement("tab-bar"));

function setBusy(busy: boolean): void {
  for (const button of actionButtons) button.disabled = busy;
}

export function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return "Excel could not complete that action.";
}

export async function refreshSelection(): Promise<void> {
  try {
    const summary = await inspectSelection();
    // -1 means the selection was too large to read (whole column/row click).
    const metric = (value: number): string =>
      value < 0 ? "—" : value.toLocaleString();
    getElement("selection-address").textContent = summary.address;
    getElement("metric-cells").textContent = metric(summary.cells);
    getElement("metric-formulas").textContent = metric(summary.formulas);
    getElement("metric-errors").textContent = metric(summary.errors);
    getElement("metric-blanks").textContent = metric(summary.blanks);
  } catch (error) {
    toast.show(errorMessage(error), "error");
  }
}

// The undo slot and the copy source are module state in excel.ts; the pane
// reads them back after every action so both rows say what they will do.
export function renderActionState(): void {
  getElement("undo-target").textContent = undoTarget() ?? "Nothing to undo yet";
  const source = copySourceLabel();
  getElement("paste-source").textContent = source
    ? `Copy source: ${source}`
    : "Mark a source, then paste it into any selection.";
}

// Every pane interaction runs through here: buttons off, toast on, busy cleared.
// renderActionState also runs after a failure: a capture may have replaced
// the undo slot already.
export const guard = makeGuard({
  setBusy,
  notify: toast.show,
  describe: (error, action) =>
    describeError(error, { host: "Excel", version: APP_VERSION }, action),
  after: refreshSelection,
  decorate: (message) =>
    lastUndoSkipped() ? `${message} (too large for undo)` : message,
  // decorate only runs on the success path, so a failed action that had
  // already skipped its undo capture would leave the flag armed for
  // whatever succeeds next; finally runs either way and drains it for good
  // (src/pane/commands.ts drains the same flag in its own finally, for the
  // ribbon's promise chain).
  finally: () => {
    lastUndoSkipped();
    renderActionState();
  },
});

// True once boot() confirms a live Excel connection; read by any tab whose
// action needs the workbook rather than the machine defaults it boots with.
let excelReady = false;

export function isExcelReady(): boolean {
  return excelReady;
}

export function setExcelReady(ready: boolean): void {
  excelReady = ready;
}

// Shared by the two "click again to confirm" delete buttons (broken names,
// unused styles): long enough to read the warning, short enough that walking
// away lands back on the safe state.
export const DELETE_CONFIRM_MS = 5_000;
