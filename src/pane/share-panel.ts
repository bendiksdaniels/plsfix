// Prepare for sharing: runs the workbook-wide cleanliness pass (every sheet to
// A1) and lists what it found - hidden sheets, external links, broken names,
// add-in-only formulas. Office.js only reaches here through ../excel.

import { prepareForSharing, type ShareResult } from "../excel";
import { type ShareIssue, summarizeShare } from "../share";
import { getElement } from "../ui/dom";
import { refreshSheets } from "./workbook-tab";

// Long lists belong in the workbook, not in a pane the reader has to scroll.
const SHARE_ROW_CAP = 20;
const SHARE_BADGES: Record<ShareIssue["kind"], string> = {
  hiddenSheet: "Hidden",
  externalLink: "External",
  addinFormula: "Needs add-in",
  brokenName: "Broken name",
  skippedSheet: "Too large",
  overlayPainted: "Overlay on",
  linkTokens: "Links",
  autocolorOnEdit: "Autocolor",
};
// Excel exposes no worksheet zoom to add-ins (pageLayout.zoom is print zoom),
// so the pane says what it cannot do rather than quietly leaving it out.
const SHARE_HINT =
  "Nothing is deleted and hidden sheets are left as they are. Zoom cannot be reset by the add-in, so it stays where you left it.";

function plural(count: number, word: string): string {
  return `${String(count)} ${count === 1 ? word : `${word}s`}`;
}

// Names how many sheets were skipped and over what cap: the row list already
// names each one (badge "Too large"), but not how many there were out of how
// many read, or the cap that decided it.
function skippedNote(result: ShareResult): string {
  const skipped = result.report.filter(
    (issue) => issue.kind === "skippedSheet",
  );
  if (skipped.length === 0) return "";
  const cap = result.sheetCap.toLocaleString();
  const names = skipped.map((issue) => issue.label).join(", ");
  return ` Scanned ${plural(result.scannedSheets, "sheet")}, ${String(skipped.length)} skipped over ${cap} cells: ${names}.`;
}

function shareRow(issue: ShareIssue): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "sheet-row";

  const label = document.createElement("span");
  label.className = "sheet-name muted";
  label.textContent = issue.label;
  label.title = issue.label;

  const badge = document.createElement("span");
  badge.className = "sheet-badge off";
  badge.textContent = SHARE_BADGES[issue.kind];

  row.append(label, badge);
  return row;
}

// A null result is the state before the first run.
export function renderShare(result: ShareResult | null): void {
  const list = getElement<HTMLDivElement>("share-report");
  list.replaceChildren();
  list.hidden = !result || result.report.length === 0;

  if (!result) {
    getElement("share-hint").textContent = SHARE_HINT;
    return;
  }

  for (const issue of result.report.slice(0, SHARE_ROW_CAP)) {
    list.append(shareRow(issue));
  }
  const rest = result.report.length - SHARE_ROW_CAP;
  if (rest > 0) {
    const more = document.createElement("p");
    more.className = "hint";
    more.textContent = `…and ${rest} more.`;
    list.append(more);
  }

  getElement("share-hint").textContent =
    `${summarizeShare(result.report, result.touchedSheets)}.${skippedNote(result)} ${SHARE_HINT}`;
}

export async function prepareShare(): Promise<string> {
  // The report lands in the pane, and the pass moves every sheet to A1: run
  // from the ribbon with the pane shut, it would rearrange the workbook and say
  // nothing. Its two ribbon siblings open the pane the same way.
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
  const result = await prepareForSharing();
  renderShare(result);
  // The pass moved the workbook, so the explorer marks the sheet it landed on.
  await refreshSheets();
  return summarizeShare(result.report, result.touchedSheets);
}
