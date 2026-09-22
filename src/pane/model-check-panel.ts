// Model check: runs the one read-only pass over the workbook and lists what it
// found, one row per finding with a jump to the cell. Holds the last report so
// the Copy button has something to hand over. Office.js only reaches here
// through ../excel; the guard and the toast come from ./shared.

import { jumpToHit, runModelCheck } from "../excel";
import {
  type Finding,
  findingLocation,
  KIND_LABELS,
  type ModelCheckReport,
  reportText,
  summarize,
} from "../model-check";
import { COPY_FAILED_MESSAGE, copyText } from "../ui/clipboard";
import { getElement } from "../ui/dom";
import { guard } from "./shared";
import { refreshSheets } from "./workbook-tab";

// Long lists belong in the copied report, not in a pane the reviewer scrolls.
const ROW_CAP = 40;
// A formula runs longer than the row; the whole of it stays in the tooltip.
const NOTE_LIMIT = 70;

const HINT =
  "One read-only pass over every sheet, hidden ones included. Nothing is painted, moved or deleted.";

// What a row with no cell to jump to says instead: where the fix lives.
const NO_JUMP: Partial<Record<Finding["kind"], string>> = {
  brokenName: "The Names section below deletes broken names.",
  unusedStyle: "The Styles section above deletes unused styles.",
};

// Null until the first run: the Copy button acts on a report, never on a guess.
let report: ModelCheckReport | null = null;

function clip(text: string): string {
  return text.length <= NOTE_LIMIT ? text : `${text.slice(0, NOTE_LIMIT)}…`;
}

function rowTitle(finding: Finding): string {
  return `${findingLocation(finding)} - ${KIND_LABELS[finding.kind]}: ${finding.note}`;
}

// A hidden sheet cannot be activated, and a workbook-level finding has no cell
// at all: both say why rather than failing silently.
async function jumpTo(finding: Finding): Promise<string> {
  if (finding.ref === null || finding.sheet === null) {
    return (
      NO_JUMP[finding.kind] ??
      `${findingLocation(finding)} has no cell to jump to.`
    );
  }
  await jumpToHit({
    kind: "cell",
    sheet: finding.sheet,
    address: finding.ref,
    text: finding.note,
  });
  // The jump moved the workbook, so the explorer above marks the sheet it
  // landed on rather than the one it left.
  await refreshSheets();
  return `Jumped to ${findingLocation(finding)}`;
}

function findingRow(finding: Finding): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.title = rowTitle(finding);

  const icon = document.createElement("span");
  icon.className = "action-icon names";
  icon.textContent = finding.kind === "formulaError" ? "!" : "•";

  const where = document.createElement("strong");
  where.textContent = findingLocation(finding);
  const what = document.createElement("small");
  what.textContent = clip(finding.note);
  const text = document.createElement("span");
  text.append(where, what);

  const badge = document.createElement("span");
  badge.className = "sheet-badge off";
  badge.textContent = KIND_LABELS[finding.kind];

  row.append(icon, text, badge);
  // Named, so a jump that cannot land says which flow refused it.
  row.addEventListener(
    "click",
    () => void guard(() => jumpTo(finding), "model-check"),
  );
  return row;
}

function hintText(current: ModelCheckReport): string {
  const skipped =
    current.skipped.length > 0
      ? ` Too large to read: ${current.skipped.join(", ")}.`
      : "";
  return `${summarize(current)}.${skipped} ${HINT}`;
}

/** A null report is the state before the first run. */
export function renderModelCheck(): void {
  const list = getElement<HTMLDivElement>("model-check-list");
  const copy = getElement<HTMLButtonElement>("copy-model-check");
  list.replaceChildren();
  copy.disabled = report === null || report.findings.length === 0;

  if (report === null) {
    list.hidden = true;
    getElement("model-check-hint").textContent = HINT;
    return;
  }

  for (const finding of report.findings.slice(0, ROW_CAP)) {
    list.append(findingRow(finding));
  }
  const rest = report.findings.length - ROW_CAP;
  if (rest > 0) {
    const more = document.createElement("p");
    more.className = "hint";
    more.textContent = `…and ${String(rest)} more in the copied report.`;
    list.append(more);
  }
  list.hidden = report.findings.length === 0;
  getElement("model-check-hint").textContent = hintText(report);
}

export async function runCheck(): Promise<string> {
  report = await runModelCheck();
  renderModelCheck();
  return summarize(report);
}

export async function copyReport(): Promise<string> {
  if (report === null) return "Run the model check first.";
  const copied = await copyText(reportText(report));
  if (!copied) throw new Error(COPY_FAILED_MESSAGE);
  return `Copied ${String(report.findings.length)} lines`;
}
