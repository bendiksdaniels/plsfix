// The Excel "Links" tab: export a selection - as a picture or as a table - or
// the active chart to PowerPoint, list what this workbook owns, push (by hand
// or automatically after an edit), jump back to a source, remove a link and
// hold the workspace link key.
// Office.js only reaches here through src/excel.
import {
  exportActiveChart,
  exportSelection,
  exportSelectionAsTable,
  goToSource,
  listWorkbookLinks,
  pushLinks,
  removeLink,
  touchWorkbookLinks,
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
import type { Toast } from "../ui/toast";
import { refreshChartPick, watchSheetChanges } from "./links-charts";
import { messageRow, renderWorkbookLinks } from "./links-list";
import {
  restoreToggles,
  setTogglesBusy,
  toggleAutoPush,
  toggleHighlight,
  type Toggles,
} from "./links-toggles";

export { renderWorkbookLinks };

const NO_KEY = "No link key yet.";
const KEY_UNREADABLE = "Could not read the link key on this computer.";
const NO_KEY_ERROR = "Generate a link key first (Links > Settings).";
const NO_SELECTION_ERROR = "Select a link in the list first.";
// Enough of the key to tell two apart, never enough to pair a deck with.
const KEY_EDGE = 4;

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
  // The two tick boxes and what they need, in the shape links-toggles.ts takes.
  toggles: Toggles;
  keyDisplay: HTMLElement;
  reveal: HTMLButtonElement;
  generate: HTMLButtonElement;
  buttons: HTMLButtonElement[];
  // The sheet's charts, for a chart export with nothing selected.
  chartPick: HTMLSelectElement;
  rows: WorkbookLinkRow[];
  selected: Set<string>;
  workspace: Workspace | null;
  // Why the stored key could not be read. Null covers both "read fine" and
  // "nothing stored"; those two are told apart by workspace.
  keyError: string | null;
  revealed: boolean;
}

export function installLinksTab(deps: LinksTabDeps): {
  refresh(): Promise<void>;
  /** The selection moved: the chart list follows if the sheet changed. */
  sheetChanged(): Promise<void>;
} {
  const tab = newTab(deps);
  wireBoxes(tab);
  wireActions(tab);
  void boot(tab);
  return {
    refresh: () => refresh(tab),
    sheetChanged: () => refreshChartPick(tab),
  };
}

function newTab(deps: LinksTabDeps): Tab {
  return {
    deps,
    chartPick: element(deps.root, "export-chart-pick"),
    list: element(deps.root, "workbook-links"),
    toggles: {
      autopush: element(deps.root, "links-autopush"),
      highlight: element(deps.root, "links-highlight"),
      relay: deps.relay,
      toast: deps.toast,
    },
    keyDisplay: element(deps.root, "workspace-key-display"),
    reveal: element(deps.root, "reveal-key"),
    generate: element(deps.root, "generate-key"),
    buttons: [],
    rows: [],
    selected: new Set(),
    workspace: null,
    keyError: null,
    revealed: false,
  };
}

// The tick boxes report through the guard like every button; "Reveal" is local
// and instant - it touches neither Office nor the store - so it stays out of
// the guard and out of the busy state.
function wireBoxes(tab: Tab): void {
  tab.reveal.addEventListener("click", () => {
    tab.revealed = !tab.revealed;
    renderKey(tab);
  });

  tab.toggles.autopush.addEventListener("change", () => {
    void guarded(tab, "links-autopush", () => toggleAutoPush(tab.toggles));
  });

  tab.toggles.highlight.addEventListener("change", () => {
    void guarded(tab, "links-highlight", () => toggleHighlight(tab.toggles));
  });
}

function wireActions(tab: Tab): void {
  wire(tab, "export-selection", () => exportRange(tab, false));
  wire(tab, "export-table", () => exportRange(tab, true));
  wire(tab, "export-chart", () => exportChart(tab));
  wire(tab, "go-to-source", () => jumpToSource(tab));
  wire(tab, "remove-link", () => removeSelected(tab));
  wire(tab, "generate-key", () => generateKey(tab));
  wire(tab, "copy-key", () => copyKey(tab));
  wire(tab, "forget-key", () => forgetKey(tab));
  wirePush(tab, "push-selected", false);
  wirePush(tab, "push-all", true);

  // Links are added and sources deleted without the pane hearing about it, so
  // the list is read again whenever the tab comes into view - and a key read
  // that failed gets another go before "Generate" is offered back.
  element(tab.deps.root, "tab-links").addEventListener("click", () => {
    void reload(tab);
  });
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

async function boot(tab: Tab): Promise<void> {
  await loadKey(tab);
  await refresh(tab);
  // Best effort: a failed touch changes nothing the user can see, and the next
  // boot tries again. Never a toast on boot.
  try {
    await touchWorkbookLinks(tab.deps.relay);
  } catch {
    // The relay is out of reach; the links keep the TTL their last push gave.
  }
  watchSheetChanges(tab);
  // Both boxes are told by the workbook, never by what they last showed. The
  // refresh above tells the same story in the table.
  await restoreToggles(tab.toggles);
}

async function reload(tab: Tab): Promise<void> {
  if (tab.keyError !== null) await loadKey(tab);
  await refresh(tab);
}

// A key that cannot be READ is not a workbook without one: answering a storage
// failure with "No link key yet." invites the modeller to generate a new key,
// which unpairs every deck holding the old one. The reason is shown instead,
// and "Generate" stays off until a read succeeds.
async function loadKey(tab: Tab): Promise<void> {
  try {
    tab.workspace = await loadWorkspace(tab.deps.keyStore);
    tab.keyError = null;
  } catch (error) {
    tab.workspace = null;
    tab.keyError = error instanceof Error ? error.message : String(error);
    tab.deps.toast.show(KEY_UNREADABLE, "error", tab.keyError);
  }
  renderKey(tab);
}

// Never rejects: every action ends with a refresh, and a list that cannot be
// read says so in the table rather than replacing the action's own toast.
async function refresh(tab: Tab): Promise<void> {
  await refreshChartPick(tab);
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
  const key = tab.workspace?.exportKey ?? null;
  tab.keyDisplay.textContent = keyText(tab, key);
  tab.reveal.textContent = tab.revealed ? "Hide" : "Reveal";
  tab.reveal.disabled = key === null;
  applyKeyState(tab);
}

// The key is the secret itself - it opens every picture this workbook pushes -
// so the panel shows only enough to tell two keys apart and leaves "Copy" as
// the route to the whole value.
function keyText(tab: Tab, key: string | null): string {
  if (tab.keyError !== null) return KEY_UNREADABLE;
  if (key === null) return NO_KEY;
  if (tab.revealed) return key;
  return `${key.slice(0, KEY_EDGE)}…${key.slice(-KEY_EDGE)}`;
}

function applyKeyState(tab: Tab): void {
  tab.generate.disabled = tab.keyError !== null;
}

function unreadable(error: unknown): string {
  const reason = error instanceof Error ? error.message : "unknown error";
  return `This workbook's links could not be read: ${reason}`;
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

async function exportRange(tab: Tab, asTable: boolean): Promise<string> {
  const send = asTable ? exportSelectionAsTable : exportSelection;
  const result = await send(requireWorkspace(tab), tab.deps.relay);
  await refresh(tab);
  return `Sent to PowerPoint: ${result.label}`;
}

async function exportChart(tab: Tab): Promise<string> {
  const pick = tab.chartPick.value || null;
  const result = await exportActiveChart(
    requireWorkspace(tab),
    tab.deps.relay,
    pick,
  );
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
  tab.keyError = null;
  tab.revealed = false;
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
  tab.revealed = false;
  renderKey(tab);
  return "Link key forgotten on this computer.";
}

function requireWorkspace(tab: Tab): Workspace {
  if (tab.keyError !== null) throw new Error(KEY_UNREADABLE);
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
  setTogglesBusy(tab.toggles, busy);
  // Busy owns every button while it runs; the key panel owns "Generate" again
  // the moment it lets go.
  if (!busy) applyKeyState(tab);
}

function element<T extends Element>(root: ParentNode, id: string): T {
  const found = root.querySelector<T>(`#${id}`);
  if (!found) throw new Error(`Missing element #${id}`);
  return found;
}
