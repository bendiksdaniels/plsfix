// The Links tab's table: one row per registry entry - a tick, the object
// (label, kind and where it points, plus a "Source missing" badge) and how
// stale the picture in a deck is - plus the full-width message row the empty
// and the unreadable state share. The internal anchor id rides along as the
// row's title (hover) only, never in the visible text. Pure DOM: handed what
// to draw and the callback to call, nothing else.

import type { WorkbookLinkRow } from "../excel";
import type { LinkKind, RegistryEntry } from "../link/model";
import { projectLabel } from "../link/project";
import { NEVER, relativeStamp } from "../ui/time";

const EMPTY_MESSAGE = "No linked objects in this workbook yet.";
const MISSING_BADGE = "Source missing";
const ROW_COLUMNS = 3;

// What each kind reads as on the row's second line.
const KIND_WORD: Record<LinkKind, string> = {
  range: "Picture",
  chart: "Chart",
  table: "Table",
  text: "Text",
};

// sourceLabel() (src/link/model.ts) appends this after a table's or a text
// link's sheet!ref address; the kind word already says it, so it is stripped
// back off rather than said twice ("Table · P&L!A10:H24", not "... table").
const KIND_SUFFIX: Partial<Record<LinkKind, string>> = {
  table: " table",
  text: " text",
};

// "Table · P&L!A10:H24": the kind in words plus where the link points, built
// from the two fields a registry entry actually keeps once a link is made -
// no raw sheet/ref survives past export, only the formatted label.
function whereItPoints(entry: RegistryEntry): string {
  const suffix = KIND_SUFFIX[entry.kind];
  const location =
    suffix !== undefined && entry.label.endsWith(suffix)
      ? entry.label.slice(0, entry.label.length - suffix.length)
      : entry.label;
  return `${KIND_WORD[entry.kind]} · ${location}`;
}

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
  const groups = new Map<string, WorkbookLinkRow[]>();
  for (const row of rows) {
    const name = projectLabel(row.entry.project);
    const group = groups.get(name) ?? [];
    group.push(row);
    groups.set(name, group);
  }
  for (const [name, group] of groups) {
    body.append(projectHeader(name));
    for (const row of group) body.append(linkRow(row, selected, onToggle));
  }
}

function projectHeader(name: string): HTMLTableRowElement {
  const tr = document.createElement("tr");
  const cell = document.createElement("td");
  cell.className = "wl-project";
  cell.colSpan = ROW_COLUMNS;
  cell.textContent = name;
  tr.append(cell);
  return tr;
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
  // The internal hidden name (PLSFIX_LINK_...) means nothing to a modeller:
  // it rides along as a hover title instead of a line of its own.
  tr.title = entry.anchor;

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
  const where = document.createElement("small");
  where.textContent = whereItPoints(row.entry);
  cell.append(label, where);

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
