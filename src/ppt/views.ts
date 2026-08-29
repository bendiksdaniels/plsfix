// The PowerPoint pane's DOM renderers: the link table body, the inbox list,
// the "Change source" candidate list and the labels they read. Pure - each
// function is handed its container, what to draw and the callback to call,
// touches nothing else, and clears the container first so a re-render can
// never leave a stale listener.

import type { InboxItem, LinkKind } from "../link/model";
import type { LinkStatus } from "../link/status";
import { NEVER, relativeStamp, relativeTime } from "../ui/time";

export interface LinkRowView {
  key: string;
  slide: number;
  label: string;
  source: string;
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

export function renderLinkRows(
  table: HTMLTableSectionElement,
  rows: LinkRowView[],
  onToggle: (key: string, selected: boolean) => void,
): void {
  table.replaceChildren(...rows.map((row) => linkRow(row, onToggle)));
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
    textCell("link-label", row.label),
    textCell("link-source", sourceMeta(row.source, row.kind)),
    statusCell(row.status),
    textCell("link-updated", relativeTime(row.pushedAt)),
  );
  return line;
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
  list.replaceChildren(...items.map((item) => inboxRow(item, onInsert)));
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
