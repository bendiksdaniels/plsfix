// The PowerPoint pane's DOM renderers: the link table body, the inbox list,
// the "Change source" candidate list and the labels they read. Pure - each
// function is handed its container, what to draw and the callback to call,
// touches nothing else, and clears the container first so a re-render can
// never leave a stale listener.

import type { InboxItem, LinkKind } from "../link/model";
import { groupByProject } from "../link/project";
import type { LinkStatus } from "../link/status";
import { NEVER, relativeStamp, relativeTime } from "../ui/time";

export interface LinkRowView {
  key: string;
  slide: number;
  label: string;
  source: string;
  project?: string;
  kind: LinkKind;
  status: LinkStatus;
  // null only for a link the relay has never held a push for; 0 is a real
  // (if ancient) time.
  pushedAt: number | null;
  selected: boolean;
}

const STATUS_LABELS: Record<LinkStatus, string> = {
  current: "Up to date",
  updateAvailable: "Update available",
  missing: "Source missing",
  wrongKey: "Wrong link key",
};

// One badge class per status so the colour is CSS's business, not the renderer's.
const STATUS_CLASSES: Record<LinkStatus, string> = {
  current: "badge current",
  updateAvailable: "badge update",
  missing: "badge missing",
  wrongKey: "badge wrong",
};

export function statusLabel(status: LinkStatus): string {
  return STATUS_LABELS[status];
}

// The table's column count: select, slide, object, status, updated. A folder
// header spans all five.
const LINK_TABLE_COLUMNS = 5;

export function renderLinkRows(
  table: HTMLTableSectionElement,
  rows: LinkRowView[],
  onToggle: (key: string, selected: boolean) => void,
): void {
  const lines: HTMLTableRowElement[] = [];
  for (const [name, group] of groupByProject(rows, (row) => row.project)) {
    lines.push(folderHeader(name, group.length));
    for (const row of group) lines.push(linkRow(row, onToggle));
  }
  table.replaceChildren(...lines);
}

// No checkbox, no data-key: a folder header is not a link row, just the
// project name and how many of the rows shown belong to it.
function folderHeader(name: string, count: number): HTMLTableRowElement {
  const line = document.createElement("tr");
  line.className = "link-folder";
  const cell = document.createElement("td");
  cell.colSpan = LINK_TABLE_COLUMNS;
  cell.textContent = `${name} · ${count}`;
  line.append(cell);
  return line;
}

function linkRow(
  row: LinkRowView,
  onToggle: (key: string, selected: boolean) => void,
): HTMLTableRowElement {
  const line = document.createElement("tr");
  line.dataset.key = row.key;
  line.append(
    selectCell(row, onToggle),
    textCell("link-slide", String(row.slide)),
    labelCell(row),
    statusCell(row.status),
    textCell("link-updated", relativeTime(row.pushedAt)),
  );
  return line;
}

// The object name plus, on its own line, the workbook it comes from and its
// kind (sourceMeta - the Inbox uses the same helper): the deck list's only
// place for the source now that the Source column is gone.
function labelCell(row: LinkRowView): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-label";
  const meta = sourceMeta(row.source, row.kind);
  const metaLine = document.createElement("span");
  metaLine.className = "link-meta";
  metaLine.textContent = meta;
  cell.append(document.createTextNode(row.label), metaLine);
  cell.title = `${row.label} · ${meta}`;
  return cell;
}

function selectCell(
  row: LinkRowView,
  onToggle: (key: string, selected: boolean) => void,
): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-select";
  const box = document.createElement("input");
  box.type = "checkbox";
  box.checked = row.selected;
  box.setAttribute("aria-label", `Select ${row.label}`);
  box.addEventListener("change", () => onToggle(row.key, box.checked));
  cell.append(box);
  return cell;
}

// The title carries the full text: these columns are narrow enough in a task
// pane that a workbook name is usually ellipsised.
function textCell(className: string, text: string): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = className;
  cell.textContent = text;
  cell.title = text;
  return cell;
}

function statusCell(status: LinkStatus): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-status";
  const badge = document.createElement("span");
  badge.className = STATUS_CLASSES[status];
  const label = statusLabel(status);
  badge.textContent = label;
  // Narrow panes truncate the badge text; the title tooltip (and an
  // aria-label, since the badge has no other accessible name) keep the
  // full status readable/announced.
  badge.title = label;
  badge.setAttribute("aria-label", label);
  cell.append(badge);
  return cell;
}

export function renderInbox(
  list: HTMLElement,
  items: InboxItem[],
  onInsert: (item: InboxItem) => void,
): void {
  if (items.length === 0) {
    list.replaceChildren(
      hint("Nothing waiting. Push a range or a chart from Excel."),
    );
    return;
  }
  list.replaceChildren(...groupedInbox(items, onInsert));
}

function groupedInbox(
  items: InboxItem[],
  onInsert: (item: InboxItem) => void,
): HTMLElement[] {
  const nodes: HTMLElement[] = [];
  for (const [name, group] of groupByProject(items, (item) => item.project)) {
    nodes.push(inboxHeader(name));
    for (const item of group) nodes.push(inboxRow(item, onInsert));
  }
  return nodes;
}

function inboxHeader(name: string): HTMLParagraphElement {
  const line = document.createElement("p");
  line.className = "inbox-project";
  line.textContent = name;
  return line;
}

function hint(text: string): HTMLParagraphElement {
  const line = document.createElement("p");
  line.className = "hint";
  line.textContent = text;
  return line;
}

function inboxRow(
  item: InboxItem,
  onInsert: (item: InboxItem) => void,
): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "inbox-row";

  const label = document.createElement("strong");
  label.textContent = item.label;
  const meta = document.createElement("small");
  meta.textContent = inboxMeta(item);
  const text = document.createElement("div");
  text.className = "inbox-text";
  text.append(label, meta);

  const insert = document.createElement("button");
  insert.type = "button";
  insert.className = "inbox-insert";
  insert.textContent = "Insert";
  insert.setAttribute("aria-label", `Insert ${item.label}`);
  insert.addEventListener("click", () => onInsert(item));

  row.append(text, insert);
  return row;
}

// The workbook a link came from and what it is: the same cells can be exported
// twice, once as a picture and once as a table, and the kind is what tells the
// two apart in a list.
function sourceMeta(workbook: string, kind: LinkKind): string {
  return `${workbook} · ${kind}`;
}

// That, plus how long the export has waited: the label above already names the
// sheet and range.
function inboxMeta(item: InboxItem): string {
  const age = relativeStamp(item.createdAt);
  const text = sourceMeta(item.src.workbook, item.kind);
  return age === NEVER ? text : `${text} · ${age}`;
}

// The inbox's Slide picker: "This slide" (value "", the default meaning "the
// active one") plus one option per slide in the deck, by position - the
// picker only ever learns a count, never the slides' own ids (target.ts
// resolves the chosen position back to one when an insert actually runs).
// The current pick survives a re-render when it is still in range, and falls
// back to "This slide" the moment the deck holds fewer slides than that.
export function renderSlideOptions(
  select: HTMLSelectElement,
  count: number,
): void {
  const picked = select.value;
  const options = [new Option("This slide", "")];
  for (let index = 1; index <= count; index += 1) {
    options.push(new Option(`Slide ${String(index)}`, String(index)));
  }
  select.replaceChildren(...options);
  select.value = options.some((option) => option.value === picked)
    ? picked
    : "";
}

// The "Change source" chooser: every waiting export, best match first (the
// order is candidatesFor's, not this renderer's). The workbook and the age are
// what tell two exports of the same table apart, so both are on the line.
export function renderCandidates(
  select: HTMLSelectElement,
  items: InboxItem[],
): void {
  select.replaceChildren(...items.map(candidateOption));
}

function candidateOption(item: InboxItem): HTMLOptionElement {
  const option = document.createElement("option");
  option.value = item.id;
  const text = `${item.label} · ${inboxMeta(item)}`;
  option.textContent = text;
  option.title = text;
  return option;
}
