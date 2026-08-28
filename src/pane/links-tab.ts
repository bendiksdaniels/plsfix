// The Excel "Links" tab: export a selection or the active chart to PowerPoint,
// list what this workbook owns, push, jump back to a source, remove a link and
// hold the workspace link key. Office.js only reaches here through src/excel.
import {
  exportActiveChart,
  exportSelection,
  goToSource,
  listWorkbookLinks,
  pushLinks,
  removeLink,
  type PushSummary,
  type WorkbookLinkRow,
} from "../excel";
import type { RelayApi } from "../link/relay";
import {
  createWorkspace,
  forgetWorkspace,
  loadWorkspace,
  type KeyStore,
  type Workspace,
} from "../link/workspace";
import { copyText } from "../ui/clipboard";
import type { Guard } from "../ui/guard";
import { relativeTime } from "../ui/time";
import type { Toast } from "../ui/toast";

const EMPTY_MESSAGE = "No linked objects in this workbook yet.";
const MISSING_BADGE = "Source missing";
const NO_KEY = "No link key yet.";
const NO_KEY_ERROR = "Generate a link key first (Links > Settings).";
const NO_SELECTION_ERROR = "Select a link in the list first.";
const ROW_COLUMNS = 3;

export interface LinksTabDeps {
  guard: Guard;
  toast: Toast;
  relay: RelayApi;
  keyStore: KeyStore;
  root: ParentNode;
}

// One object threaded through the actions, so each stays a small function over
// the same state instead of a closure inside a long install().
interface Tab {
  deps: LinksTabDeps;
  list: HTMLTableSectionElement;
  keyDisplay: HTMLElement;
  buttons: HTMLButtonElement[];
  rows: WorkbookLinkRow[];
  selected: Set<string>;
  workspace: Workspace | null;
}

export function installLinksTab(deps: LinksTabDeps): {
  refresh(): Promise<void>;
} {
  const tab: Tab = {
    deps,
    list: element(deps.root, "workbook-links"),
    keyDisplay: element(deps.root, "workspace-key-display"),
    buttons: [],
    rows: [],
    selected: new Set(),
    workspace: null,
  };

  wire(tab, "export-selection", () => exportRange(tab));
  wire(tab, "export-chart", () => exportChart(tab));
  wire(tab, "go-to-source", () => jumpToSource(tab));
  wire(tab, "remove-link", () => removeSelected(tab));
  wire(tab, "generate-key", () => generateKey(tab));
  wire(tab, "copy-key", () => copyKey(tab));
  wire(tab, "forget-key", () => forgetKey(tab));
  wirePush(tab, "push-selected", false);
  wirePush(tab, "push-all", true);

  // Links are added and sources deleted without the pane hearing about it, so
  // the list is read again whenever the tab comes into view.
  element(deps.root, "tab-links").addEventListener("click", () => {
    void refresh(tab);
  });

  void boot(tab);
  return { refresh: () => refresh(tab) };
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
  for (const row of rows) body.append(linkRow(row, selected, onToggle));
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

async function boot(tab: Tab): Promise<void> {
  tab.workspace = await loadWorkspace(tab.deps.keyStore).catch(() => null);
  renderKey(tab);
  await refresh(tab);
}

// Never rejects: every action ends with a refresh, and a list that cannot be
// read says so in the table rather than replacing the action's own toast.
async function refresh(tab: Tab): Promise<void> {
  try {
    tab.rows = await listWorkbookLinks();
  } catch (error) {
    tab.rows = [];
    tab.selected.clear();
    tab.list.replaceChildren(messageRow(unreadable(error)));
    return;
  }
  const live = new Set(tab.rows.map((row) => row.entry.id));
  for (const id of tab.selected) if (!live.has(id)) tab.selected.delete(id);
  renderWorkbookLinks(tab.list, tab.rows, tab.selected, (id, on) => {
    if (on) tab.selected.add(id);
    else tab.selected.delete(id);
  });
}

function renderKey(tab: Tab): void {
  tab.keyDisplay.textContent = tab.workspace?.exportKey ?? NO_KEY;
}

function unreadable(error: unknown): string {
  const reason = error instanceof Error ? error.message : "unknown error";
  return `This workbook's links could not be read: ${reason}`;
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

async function exportRange(tab: Tab): Promise<string> {
  const result = await exportSelection(requireWorkspace(tab), tab.deps.relay);
  await refresh(tab);
  return `Sent to PowerPoint: ${result.label}`;
}

async function exportChart(tab: Tab): Promise<string> {
  const result = await exportActiveChart(requireWorkspace(tab), tab.deps.relay);
  await refresh(tab);
  return `Sent to PowerPoint: ${result.label}`;
}

// pushLinks reports rather than throws, so a partial failure arrives as a
// summary: the guard toasts the counts, and the reasons are re-shown here with
// a "Copy details" button. Both happen before a paint, so only one is seen.
async function push(tab: Tab, action: string, all: boolean): Promise<void> {
  const failures: string[] = [];
  let line = "";
  await guarded(tab, action, async () => {
    const ids = all ? "all" : [...requireSelection(tab)];
    const summary = await pushLinks(ids, tab.deps.relay);
    failures.push(...summary.failures);
    line = summarize(summary);
    await refresh(tab);
    return line;
  });
  if (failures.length > 0) {
    tab.deps.toast.show(line, "error", failures.join("\n"));
  }
}

function summarize(summary: PushSummary): string {
  return `${summary.pushed} pushed, ${summary.missing} missing, ${summary.failed} failed`;
}

async function jumpToSource(tab: Tab): Promise<string> {
  const [id] = requireSelection(tab);
  const row = tab.rows.find((candidate) => candidate.entry.id === id);
  await goToSource(id);
  return `Went to ${row?.entry.label ?? "the source"}`;
}

async function removeSelected(tab: Tab): Promise<string> {
  const ids = requireSelection(tab);
  try {
    for (const id of ids) await removeLink(id, tab.deps.relay);
  } finally {
    // A failure halfway leaves some links gone: the list must show which.
    await refresh(tab);
  }
  return `Removed ${ids.length} ${ids.length === 1 ? "link" : "links"}`;
}

async function generateKey(tab: Tab): Promise<string> {
  tab.workspace = await createWorkspace(tab.deps.keyStore);
  renderKey(tab);
  return "Link key generated. Paste it in PowerPoint.";
}

async function copyKey(tab: Tab): Promise<string> {
  await copyText(requireWorkspace(tab).exportKey);
  return "Link key copied.";
}

async function forgetKey(tab: Tab): Promise<string> {
  await forgetWorkspace(tab.deps.keyStore);
  tab.workspace = null;
  renderKey(tab);
  return "Link key forgotten on this computer.";
}

function requireWorkspace(tab: Tab): Workspace {
  if (tab.workspace === null) throw new Error(NO_KEY_ERROR);
  return tab.workspace;
}

// Read back through the rendered rows, so a tick left over from a link that no
// longer exists cannot reach the adapter.
function requireSelection(tab: Tab): [string, ...string[]] {
  const [first, ...rest] = tab.rows
    .filter((row) => tab.selected.has(row.entry.id))
    .map((row) => row.entry.id);
  if (first === undefined) throw new Error(NO_SELECTION_ERROR);
  return [first, ...rest];
}

// ---------------------------------------------------------------------------
// Wiring
// ---------------------------------------------------------------------------

function wire(tab: Tab, id: string, run: () => Promise<string>): void {
  listen(tab, id, () => guarded(tab, id, run));
}

function wirePush(tab: Tab, id: string, all: boolean): void {
  listen(tab, id, () => push(tab, id, all));
}

function listen(tab: Tab, id: string, run: () => Promise<void>): void {
  const button = element<HTMLButtonElement>(tab.deps.root, id);
  tab.buttons.push(button);
  button.addEventListener("click", () => {
    void run();
  });
}

// The shared guard disables main.ts's [data-action] buttons, which these are
// not: a second click during a slow push would start a second one.
async function guarded(
  tab: Tab,
  action: string,
  run: () => Promise<string>,
): Promise<void> {
  await tab.deps.guard(async () => {
    setBusy(tab, true);
    try {
      return await run();
    } finally {
      setBusy(tab, false);
    }
  }, action);
}

function setBusy(tab: Tab, busy: boolean): void {
  for (const button of tab.buttons) button.disabled = busy;
}

function element<T extends Element>(root: ParentNode, id: string): T {
  const found = root.querySelector<T>(`#${id}`);
  if (!found) throw new Error(`Missing element #${id}`);
  return found;
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

// One full-width cell, used for both the empty list and a list that could not
// be read: the table itself says why rather than a toast that fades.
function messageRow(text: string): HTMLTableRowElement {
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
  if (lastPushedAt === null) return "never";
  const seconds = Date.parse(lastPushedAt) / 1000;
  if (!Number.isFinite(seconds)) return "never";
  return `Pushed ${relativeTime(seconds)}`;
}
