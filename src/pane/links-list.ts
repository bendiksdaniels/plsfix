// The Links tab's table: one row per registry entry - a tick, the object
// (label, anchor and a "Source missing" badge) and how stale the picture in a
// deck is - plus the full-width message row the empty and the unreadable state
// share. Pure DOM: handed what to draw and the callback to call, nothing else.

import type { WorkbookLinkRow } from "../excel";
import { NEVER, relativeStamp } from "../ui/time";

const EMPTY_MESSAGE = "No linked objects in this workbook yet.";
const MISSING_BADGE = "Source missing";
const ROW_COLUMNS = 3;

export function renderWorkbookLinks(
  body: HTMLTableSectionElement,
  rows: WorkbookLinkRow[],
  selected: Set<string>,
  onToggle: (id: string, on: boolean) => void,
): void {
  body.replaceChildren();
  if (rows.length === 0) {
    body.append(messageRow(EMPTY_MESSAGE));
    return;
  }
  for (const row of rows) body.append(linkRow(row, selected, onToggle));
}

// One full-width cell, used for both the empty list and a list that could not
// be read: the table itself says why rather than a toast that fades.
export function messageRow(text: string): HTMLTableRowElement {
  const tr = document.createElement("tr");
  const cell = document.createElement("td");
  cell.className = "wl-empty";
  cell.colSpan = ROW_COLUMNS;
  cell.textContent = text;
  tr.append(cell);
  return tr;
}

function linkRow(
  row: WorkbookLinkRow,
  selected: Set<string>,
  onToggle: (id: string, on: boolean) => void,
): HTMLTableRowElement {
  const { entry } = row;
  const tr = document.createElement("tr");
  tr.dataset.linkId = entry.id;

  const pick = document.createElement("td");
  pick.className = "wl-pick";
  const box = document.createElement("input");
  box.type = "checkbox";
  box.checked = selected.has(entry.id);
  box.setAttribute("aria-label", `Select ${entry.label}`);
  box.addEventListener("change", () => {
    onToggle(entry.id, box.checked);
  });
  pick.append(box);

  const pushed = document.createElement("td");
  pushed.className = "wl-pushed";
  pushed.textContent = pushedLabel(entry.lastPushedAt);

  tr.append(pick, objectCell(row), pushed);
  return tr;
}

function objectCell(row: WorkbookLinkRow): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "wl-object";

  const label = document.createElement("strong");
  label.textContent = row.entry.label;
  const anchor = document.createElement("small");
  anchor.textContent = row.entry.anchor;
  cell.append(label, anchor);

  if (row.source === "missing") {
    const badge = document.createElement("span");
    badge.className = "wl-badge";
    badge.textContent = MISSING_BADGE;
    cell.append(badge);
  }
  return cell;
}

// How stale the picture in a deck is, not when it was made, so the list reads
// as an age. An unparseable stamp is treated as no push at all.
function pushedLabel(lastPushedAt: string | null): string {
  const age = relativeStamp(lastPushedAt);
  return age === NEVER ? age : `Pushed ${age}`;
}
