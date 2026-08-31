// Style scrubber: scans the workbook's custom cell styles for ones no cell
// wears any more and deletes them. The "click again to confirm" arm on the
// delete button lapses on its own. Office.js only reaches here through
// ../excel.

import { deleteUnusedStyles, listUnusedStyles, type StyleScan } from "../excel";
import { getElement } from "../ui/dom";
import { DELETE_CONFIRM_MS, tabs } from "./shared";

let styleScan: StyleScan | null = null;
let stylesArmed = false;
let stylesTimer: number | undefined;

export function isStylesArmed(): boolean {
  return stylesArmed;
}

function unusedCount(): string {
  const count = styleScan?.unused.length ?? 0;
  return `${count} unused ${count === 1 ? "style" : "styles"}`;
}

function stylesDeleteLabel(): string {
  return `Delete ${unusedCount()}`;
}

export function disarmStyles(): void {
  window.clearTimeout(stylesTimer);
  stylesArmed = false;
  const button = getElement<HTMLButtonElement>("styles-delete");
  button.classList.remove("armed");
  button.textContent = stylesDeleteLabel();
}

// Deleting a style restyles every cell wearing it and no undo brings it back,
// so the first click only arms the button and the arming lapses on its own.
export function armStyles(): void {
  const button = getElement<HTMLButtonElement>("styles-delete");
  stylesArmed = true;
  button.classList.add("armed");
  button.textContent = "Click again to confirm";
  stylesTimer = window.setTimeout(disarmStyles, DELETE_CONFIRM_MS);
}

// A sheet too large to read could be wearing any of these styles, so the count
// is stated as incomplete and the delete stays out of reach until it is not.
function stylesSummary(scan: StyleScan): string {
  const table = ` of ${scan.total} in the workbook`;
  const skipped =
    scan.skippedSheets.length > 0
      ? ` Some sheets were too large to scan: ${scan.skippedSheets.join(", ")}.`
      : "";
  if (scan.unused.length === 0) {
    return `No unused custom styles${table}.${skipped}`;
  }
  return `${unusedCount()}${table}.${skipped}`;
}

function styleRow(name: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "sheet-row";
  // Muted: a style name is a label, not somewhere to click through to.
  const label = document.createElement("span");
  label.className = "sheet-name muted";
  label.textContent = name;
  row.append(label);
  return row;
}

export function renderStyles(): void {
  const list = getElement<HTMLDivElement>("styles-list");
  const button = getElement<HTMLButtonElement>("styles-delete");
  list.replaceChildren();

  if (!styleScan) {
    getElement("styles-result").textContent = "Not scanned yet.";
    list.hidden = true;
    button.hidden = true;
    return;
  }

  for (const name of styleScan.unused) list.append(styleRow(name));
  list.hidden = styleScan.unused.length === 0;
  getElement("styles-result").textContent = stylesSummary(styleScan);

  disarmStyles();
  button.hidden = styleScan.unused.length === 0;
  button.disabled = styleScan.skippedSheets.length > 0;
}

export async function scanStyles(): Promise<string> {
  styleScan = await listUnusedStyles();
  renderStyles();
  return stylesSummary(styleScan);
}

// Rescanned afterwards rather than assumed: the style table is what shrank, and
// the pane says what is left of it.
export async function deleteStyles(): Promise<string> {
  const removed = await deleteUnusedStyles(styleScan?.unused ?? []);
  styleScan = await listUnusedStyles();
  renderStyles();
  return `Deleted ${removed} unused ${removed === 1 ? "style" : "styles"}`;
}

// The list lives in the pane, so the command opens it on the Workbook tab and
// leaves the result on screen rather than reporting a number and forgetting it.
export async function focusStyles(): Promise<string> {
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
  tabs.activate("tab-workbook");
  return scanStyles();
}
