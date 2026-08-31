// Super Find: searches values, defined names, sheet names and comments across
// the workbook and jumps to a hit. Office.js only reaches here through
// ../excel; the guard, toast and tab switcher are shared through ./shared.

import {
  findInWorkbook,
  type FindHit,
  type FindResult,
  jumpToHit,
} from "../excel";
import { FIND_HIT_CAP } from "../find";
import { getElement } from "../ui/dom";
import { guard, tabs } from "./shared";
import { refreshSheets } from "./workbook-tab";

const FIND_TEXT_LIMIT = 90;
const FIND_ICONS: Record<FindHit["kind"], string> = {
  cell: "▤",
  name: "⌗",
  sheet: "☰",
  comment: "❝",
};

function hitLabel(hit: FindHit): string {
  if (hit.kind === "name") return `Name · ${hit.address}`;
  if (hit.kind === "comment") return `Comment · ${hit.sheet}!${hit.address}`;
  return `${hit.sheet}!${hit.address}`;
}

// A long label or a formula would push the row out of the pane; the whole text
// stays in the tooltip.
function clipText(text: string): string {
  if (text.length <= FIND_TEXT_LIMIT) return text;
  return `${text.slice(0, FIND_TEXT_LIMIT)}…`;
}

async function jumpTo(hit: FindHit): Promise<string> {
  await jumpToHit(hit);
  // The jump moved the workbook, so the explorer above marks the sheet it
  // landed on rather than the one it left.
  await refreshSheets();
  return `Jumped to ${hitLabel(hit)}`;
}

function findRow(hit: FindHit): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.title = `Go to ${hitLabel(hit)}: ${hit.text}`;

  const icon = document.createElement("span");
  icon.className = "action-icon names";
  icon.textContent = FIND_ICONS[hit.kind];

  const where = document.createElement("strong");
  where.textContent = hitLabel(hit);
  const what = document.createElement("small");
  what.textContent = clipText(hit.text);
  const text = document.createElement("span");
  text.append(where, what);

  row.append(icon, text);
  // Named, so a jump that cannot land says which flow refused it.
  row.addEventListener("click", () => void guard(() => jumpTo(hit), "find"));
  return row;
}

function findSummary(result: FindResult): string {
  const count = result.hits.length;
  const capped = count >= FIND_HIT_CAP ? ` (first ${FIND_HIT_CAP})` : "";
  const skipped =
    result.skippedSheets.length > 0
      ? ` Too large to search: ${result.skippedSheets.join(", ")}.`
      : "";
  // An old host has no comment collection at all, so "no matches" would read
  // as "nothing was written there" rather than "nobody looked".
  const gated = result.commentsSkipped ? " Comments need Excel 365." : "";
  if (count === 0) return `No matches.${skipped}${gated}`;
  return `${count} ${count === 1 ? "hit" : "hits"}${capped}.${skipped}${gated}`;
}

// Rows close over the hit they jump to; drop them before rebuilding. A null
// result is the state before the first search.
export function renderFind(result: FindResult | null): void {
  const list = getElement<HTMLDivElement>("find-results");
  list.replaceChildren();
  // An empty list still costs a row gap under the button, so it goes away.
  list.hidden = !result || result.hits.length === 0;

  if (!result) {
    getElement("find-hint").textContent =
      "Searches values, workbook-level defined names, sheet names and comments on every sheet.";
    return;
  }
  for (const hit of result.hits) list.append(findRow(hit));
  getElement("find-hint").textContent = findSummary(result);
}

export async function runFind(): Promise<string> {
  const query = getElement<HTMLInputElement>("find-query").value.trim();
  // An empty box is not a failure: it reads back as a plain note rather than a
  // red toast with a "Copy details" button behind it.
  if (query === "") {
    renderFind(null);
    return "Type something to find first.";
  }

  const result = await findInWorkbook(query, {
    matchCase: getElement<HTMLInputElement>("find-case").checked,
    inFormulas: getElement<HTMLInputElement>("find-formulas").checked,
    inComments: getElement<HTMLInputElement>("find-comments").checked,
  });
  renderFind(result);
  return findSummary(result);
}

// The results live in the pane, so the shortcut opens it and puts the caret in
// the box rather than repeating a query the modeller cannot see.
export async function focusFind(): Promise<string> {
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
  tabs.activate("tab-workbook");
  await refreshSheets();

  const input = getElement<HTMLInputElement>("find-query");
  input.focus();
  input.select();
  return "Find ready";
}
