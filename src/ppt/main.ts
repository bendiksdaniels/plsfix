// PowerPoint pane: boot (Office.js, the PowerPoint and PowerPointApi 1.5
// gates, toast, error reporting, tabs, version), the pane's state - the links
// this deck holds, which are ticked, and the paired workspace - and one guarded
// handler per button. Drawing is views.ts, selection maths actions.ts, the
// "Change source" picker chooser.ts, and deck/relay calls links.ts / host.ts.

import "../styles.css";
import type { InboxItem } from "../link/model";
import { relayBaseUrl, RelayClient } from "../link/relay";
import {
  forgetWorkspace,
  importWorkspace,
  loadWorkspace,
  officeKeyStore,
  type Workspace,
} from "../link/workspace";
import { getElement } from "../ui/dom";
import { installFirstRun } from "../ui/first-run";
import { makeGuard } from "../ui/guard";
import { installHelp } from "../ui/help";
import { describeError, installErrorReporting } from "../ui/report";
import { installTabs } from "../ui/tabs";
import { createToast } from "../ui/toast";
import { formatVersion } from "../ui/version";
import {
  filterRows,
  linkSlides,
  linkSources,
  pruneSelection,
  requireSelection,
  selectedRows,
  slideRows,
  toRowViews,
  updateDetails,
} from "./actions";
import { installChangeSource } from "./chooser";
import { registerCommands } from "./commands";
import { activeSlideId, breakLink, goToSlide } from "./host";
import { latestInboxItem } from "./inbox-queue";
import {
  insertFromInbox,
  insertNote,
  listInbox,
  listLinks,
  summarize,
  updateLinks,
  type LinkRow,
} from "./links";
import { revertLinks, summarizeRevert } from "./revert";
import {
  alignSelected,
  applyObjectStyle,
  captureObjectStyle,
  distributeSelected,
  matchSelectedSize,
  selectSimilar,
  swapSelected,
} from "./object-tools";
import { createPaneDetails } from "./pane-details";
import { readInsertTarget, refreshSlideOptions } from "./target";
import { renderInbox, renderLinkRows } from "./views";

const APP_VERSION = formatVersion(__APP_VERSION__);
const REPORT_CONTEXT = { host: "PowerPoint", version: APP_VERSION };
const NOT_PAIRED =
  "Not paired: paste the link key from Excel > Links > Link key";
const PAIR_FIRST = "Paste the link key in Settings";
const STILL_BUSY = "Wait for the last action to finish.";

const connectionStatus = getElement<HTMLSpanElement>("connection-status");
const linkRowsBody = getElement<HTMLTableSectionElement>("link-rows");
const linksEmpty = getElement("links-empty");
const linksFilteredEmpty = getElement("links-filtered-empty");
const linkSearch = getElement<HTMLInputElement>("link-search");
const linkStatusFilter = getElement<HTMLSelectElement>("link-status-filter");
const linkSourceFilter = getElement<HTMLSelectElement>("link-source-filter");
const linkSlideFilter = getElement<HTMLSelectElement>("link-slide-filter");
const objectAlignMode = getElement<HTMLSelectElement>("object-align-mode");
const objectDistributeAxis = getElement<HTMLSelectElement>(
  "object-distribute-axis",
);
const insertSlideSelect = getElement<HTMLSelectElement>("insert-slide");
const insertWhereSelect = getElement<HTMLSelectElement>("insert-where");
const inboxList = getElement("inbox-list");
const inboxUnpaired = getElement("inbox-unpaired");
const workspaceState = getElement("workspace-state");
const workspaceKey = getElement<HTMLInputElement>("workspace-key");
const toast = createToast(getElement("toast"));

const relay = new RelayClient(relayBaseUrl(document.baseURI));
const keyStore = officeKeyStore();

let rows: LinkRow[] = [];
let inboxItems: InboxItem[] = [];
let workspace: Workspace | null = null;
const selected = new Set<string>();

getElement("app-version").textContent = APP_VERSION;

// Installed first so a throw during the rest of boot is still reported.
installErrorReporting(REPORT_CONTEXT, (message, details) =>
  toast.show(message, "error", details),
);
const tabs = installTabs(getElement("tab-bar"));
// Ribbon commands share this runtime, so they register with the page and
// toast into it whether or not the pane is showing.
registerCommands({
  notify: (message, kind, lines) => toast.show(message, kind, lines),
  context: REPORT_CONTEXT,
  showTools: async () => {
    await Office.addin?.showAsTaskpane();
    tabs.activate("tab-tools");
  },
  // Through the pane's own guard: one action at a time, ribbon included.
  run: (work, action) => guard(work, action),
});
// The "?" on every section heading, added once the markup is in place.
installHelp(document);
// The Inbox tab's "New here?" card, dismissed for good on this machine.
installFirstRun(document, "plsfix.firstRun.ppt.v1", "first-run-ppt");

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

function renderLinks(): void {
  syncLinkFilterOptions();
  const filtered = filterRows(rows, {
    query: linkSearch.value,
    status: linkStatusFilter.value as Parameters<
      typeof filterRows
    >[1]["status"],
    source: linkSourceFilter.value,
    slide:
      linkSlideFilter.value === "all" ? "all" : Number(linkSlideFilter.value),
  });
  renderLinkRows(linkRowsBody, toRowViews(filtered, selected), toggleSelection);
  linksEmpty.hidden = rows.length > 0;
  linksFilteredEmpty.hidden = rows.length === 0 || filtered.length > 0;
  syncChangeSource();
}

function replaceFilterOptions(
  select: HTMLSelectElement,
  allLabel: string,
  values: (string | number)[],
  label: (value: string | number) => string,
): void {
  const selectedValue = select.value;
  const options = [new Option(allLabel, "all")];
  for (const value of values) {
    options.push(new Option(label(value), String(value)));
  }
  select.replaceChildren(...options);
  select.value = options.some((option) => option.value === selectedValue)
    ? selectedValue
    : "all";
}

function syncLinkFilterOptions(): void {
  replaceFilterOptions(
    linkSourceFilter,
    "All sources",
    linkSources(rows),
    String,
  );
  replaceFilterOptions(
    linkSlideFilter,
    "All slides",
    linkSlides(rows),
    (value) => `Slide ${String(value)}`,
  );
}

function toggleSelection(key: string, isSelected: boolean): void {
  if (isSelected) selected.add(key);
  else selected.delete(key);
  syncChangeSource();
}

// The list is hidden rather than emptied when unpaired, and its children are
// replaced either way: a stale Insert button must never survive a re-render.
function renderInboxView(): void {
  renderInbox(inboxList, inboxItems, (item) => {
    act(() => insertItem(item), "insert-link");
  });
  // An insert redraws this list halfway through - the row it consumed has to
  // leave - so the Insert buttons that redraw creates are ones setBusy has
  // never seen. Without this a second insert is clickable mid-flight.
  for (const insert of inboxList.querySelectorAll("button")) {
    insert.disabled = isBusy;
  }
  inboxList.hidden = workspace === null;
  inboxUnpaired.hidden = workspace !== null;
}

function renderPairing(): void {
  workspaceState.textContent = workspace === null ? NOT_PAIRED : "Paired";
}

// ---------------------------------------------------------------------------
// The guard every button runs through
// ---------------------------------------------------------------------------

// makeGuard's success path notifies with a message only, so an action with
// per-link lines to show stages them here; the guard reads and clears them.
const details = createPaneDetails(REPORT_CONTEXT);

// What setBusy last wrote: a list redrawn while an action is still running
// has to come back as disabled as the buttons it replaced.
let isBusy = false;

function setBusy(busy: boolean): void {
  isBusy = busy;
  const buttons =
    document.querySelectorAll<HTMLButtonElement>(".app-shell button");
  for (const button of buttons) button.disabled = busy;
  // Blanket re-enabling would undo the one button with a rule of its own.
  if (!busy) syncChangeSource();
}

const guard = makeGuard({
  setBusy,
  // A ribbon command whose whole answer is the pane coming up returns "".
  notify: (message, kind, lines) => {
    if (message !== "") toast.show(message, kind, lines ?? details.value);
  },
  describe: (error, action) => describeError(error, REPORT_CONTEXT, action),
  finally: () => {
    details.set(undefined);
  },
  busyMessage: STILL_BUSY,
});

// Set once Office.onReady has confirmed PowerPoint and PowerPointApi 1.5.
let ready = false;

// Every button is live from the first paint, so a click during a cold boot - or
// in a deck on a PowerPoint the pane rejected - would reach PowerPoint.run and
// toast a raw "Cannot read properties of undefined". Every handler starts here
// instead, and the connection badge is not the only thing that says so.
function act(run: () => Promise<string>, action: string): void {
  void guard(async () => {
    if (!ready) throw new Error("PowerPoint is not connected.");
    return run();
  }, action);
}

// ---------------------------------------------------------------------------
// Links
// ---------------------------------------------------------------------------

async function reloadLinks(): Promise<void> {
  rows = await listLinks(relay);
  pruneSelection(rows, selected);
  renderLinks();
}

// What an action reports is the point of it, so a rescan that fails afterwards
// leaves the list as it was and says so in the details, not in place of it.
async function refreshQuietly(): Promise<void> {
  try {
    await reloadLinks();
  } catch (error) {
    details.addFailure(error, "The list was not refreshed");
  }
}

async function refreshLinks(): Promise<string> {
  await reloadLinks();
  if (rows.length === 0) return "No linked objects in this deck.";
  return `${String(rows.length)} linked ${rows.length === 1 ? "object" : "objects"}.`;
}

async function updateRows(subset: LinkRow[]): Promise<string> {
  const summary = await updateLinks(subset, relay);
  details.set(updateDetails(summary));
  await refreshQuietly();
  return summarize(summary);
}

// "Update this slide" reads PowerPoint's own selection, not a tick: the pane
// cannot see it any other way, so a slide with nothing selected or nothing
// linked on it is reported rather than silently falling back to a tick.
async function updateSlide(): Promise<string> {
  const slideId = await activeSlideId();
  if (slideId === null) throw new Error("Select a slide first.");
  const subset = slideRows(rows, slideId);
  if (subset.length === 0) throw new Error("No links on this slide");
  return updateRows(subset);
}

// PowerPoint cannot undo what the pane wrote, and the relay keeps only the
// revision before the current one, so this goes back exactly one step - on the
// ticked rows, because a whole-deck revert is not something to reach by
// accident.
async function revertSelected(): Promise<string> {
  const subset = requireSelection(
    selectedRows(rows, selected),
    "Tick the rows to revert.",
  );
  const summary = await revertLinks(subset, relay);
  details.set(
    summary.failures.length > 0 ? summary.failures.join("\n") : undefined,
  );
  await refreshQuietly();
  return summarizeRevert(summary);
}

async function breakSelected(): Promise<string> {
  const subset = requireSelection(selectedRows(rows, selected));
  try {
    for (const row of subset) await breakLink(row.found);
  } finally {
    // Whatever happened, the list must show what the deck now holds.
    await refreshQuietly();
  }
  const count = `${String(subset.length)} link${subset.length === 1 ? "" : "s"}`;
  return `${count} broken. The picture stays on the slide.`;
}

async function goToSelectedSlide(): Promise<string> {
  const first = requireSelection(selectedRows(rows, selected))[0]!;
  await goToSlide(first.found.slideId);
  return `Slide ${String(first.found.slideIndex + 1)}.`;
}

// ---------------------------------------------------------------------------
// Inbox and pairing
// ---------------------------------------------------------------------------

function requireWorkspace(): Workspace {
  if (workspace === null) throw new Error(PAIR_FIRST);
  return workspace;
}

async function refreshInbox(): Promise<string> {
  const ws = requireWorkspace();
  await refreshSlideOptions(insertSlideSelect, (error) => {
    details.addFailure(error, "The slide list was not refreshed");
  });
  inboxItems = await listInbox(ws, relay);
  renderInboxView();
  if (inboxItems.length === 0) return "Nothing waiting from Excel.";
  return `${String(inboxItems.length)} waiting to insert.`;
}

async function insertItem(item: InboxItem): Promise<string> {
  const ws = requireWorkspace();
  const target = await readInsertTarget(insertSlideSelect, insertWhereSelect);
  const placed = await insertFromInbox(item, ws, relay, target);
  // The relay copy is gone, so the item leaves the list without a second call.
  inboxItems = inboxItems.filter((waiting) => waiting.id !== item.id);
  renderInboxView();
  await refreshQuietly();
  return `Inserted ${item.label}.${insertNote(placed)}`;
}

// The normal cross-app flow: Excel exports, then PowerPoint puts the freshest
// waiting object on the active slide and tags it in the same action. Read the
// relay here rather than trusting the rendered Inbox, which may be seconds old.
async function pasteLatestLinked(): Promise<string> {
  const ws = requireWorkspace();
  inboxItems = await listInbox(ws, relay);
  const item = latestInboxItem(inboxItems);
  if (item === null) {
    renderInboxView();
    throw new Error(
      "Nothing waiting from Excel. Export an object there first.",
    );
  }
  return insertItem(item);
}

// The "Change source" picker owns its own three buttons; the pane hands it the
// state it must read and the reads that follow a successful change.
const syncChangeSource = installChangeSource({
  act,
  relay,
  rows: () => selectedRows(rows, selected),
  inbox: () => inboxItems,
  workspace: requireWorkspace,
  note: (line) => {
    details.add(line);
  },
  after: async () => {
    await inboxQuietly();
    await refreshQuietly();
  },
});

async function saveKey(): Promise<string> {
  const key = workspaceKey.value.trim();
  if (key === "") throw new Error("Paste the link key from Excel first.");
  workspace = await pairWith(key);
  // The key is the secret itself: it is stored, never echoed back.
  workspaceKey.value = "";
  // The exports on screen were sealed with the key that was there before, and
  // their Insert buttons still work: they go with it, whether or not the read
  // below reaches the relay.
  inboxItems = [];
  // Paired is drawn before the inbox is read, so a relay that is down leaves an
  // empty list rather than a pane that still claims to be unpaired.
  renderPairing();
  renderInboxView();
  await inboxQuietly();
  return "Paired with Excel.";
}

// importWorkspace names itself in its message; a mistyped key is the user's
// business, a storage failure is the pane's, so only the first is reworded.
async function pairWith(key: string): Promise<Workspace> {
  try {
    return await importWorkspace(keyStore, key);
  } catch (error) {
    if (error instanceof Error && error.message.includes("invalid link key")) {
      throw new Error("That is not a link key. Copy it again from Excel.");
    }
    throw error;
  }
}

async function forgetKey(): Promise<string> {
  await forgetWorkspace(keyStore);
  workspace = null;
  inboxItems = [];
  renderPairing();
  renderInboxView();
  return "Link key forgotten. The links already in this deck still update.";
}

// The inbox is a courtesy after pairing: a relay that is down must not turn a
// saved key into a failure.
async function inboxQuietly(): Promise<void> {
  if (workspace === null) return;
  try {
    await refreshInbox();
  } catch (error) {
    details.addFailure(error, "The inbox was not read");
  }
}

// ---------------------------------------------------------------------------
// Wiring and boot
// ---------------------------------------------------------------------------

const BUTTON_ACTIONS: Record<string, () => Promise<string>> = {
  "refresh-links": refreshLinks,
  "update-selected": () =>
    updateRows(requireSelection(selectedRows(rows, selected))),
  "update-slide": updateSlide,
  "update-all": () => updateRows(rows),
  "revert-selected": revertSelected,
  "break-selected": breakSelected,
  "go-to-slide": goToSelectedSlide,
  "refresh-inbox": refreshInbox,
  "paste-latest-linked": pasteLatestLinked,
  "save-key": saveKey,
  "forget-key": forgetKey,
  "align-objects": () =>
    alignSelected(objectAlignMode.value as Parameters<typeof alignSelected>[0]),
  "distribute-objects": () =>
    distributeSelected(
      objectDistributeAxis.value as Parameters<typeof distributeSelected>[0],
    ),
  "match-size": matchSelectedSize,
  "select-similar": selectSimilar,
  "swap-objects": swapSelected,
  "capture-object-style": captureObjectStyle,
  "apply-object-style": applyObjectStyle,
};

for (const [id, run] of Object.entries(BUTTON_ACTIONS)) {
  getElement<HTMLButtonElement>(id).addEventListener("click", () => {
    act(run, id);
  });
}

workspaceKey.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  act(saveKey, "save-key");
});

linkSearch.addEventListener("input", renderLinks);
linkStatusFilter.addEventListener("change", renderLinks);
linkSourceFilter.addEventListener("change", renderLinks);
linkSlideFilter.addEventListener("change", renderLinks);

renderLinks();
renderInboxView();
renderPairing();

// Boot never blanks the pane: each step reports its own failure to the toast
// and the next one still runs, so a relay that is down leaves empty lists and
// a working pane rather than nothing at all.
async function bootStep(
  run: () => Promise<unknown>,
  action: string,
): Promise<void> {
  try {
    await run();
  } catch (error) {
    const { message, details } = describeError(error, REPORT_CONTEXT, action);
    toast.show(message, "error", details);
  }
}

async function loadPairing(): Promise<void> {
  workspace = await loadWorkspace(keyStore);
  renderPairing();
  renderInboxView();
}

function showConnection(text: string, state: string): void {
  connectionStatus.textContent = text;
  connectionStatus.className = `connection ${state}`;
}

Office.onReady(async ({ host }) => {
  if (host !== Office.HostType.PowerPoint) {
    showConnection("PowerPoint required", "error");
    return;
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    showConnection("PowerPoint 2021 / Microsoft 365 required", "error");
    return;
  }
  showConnection("PowerPoint connected", "ready");
  ready = true;

  await bootStep(loadPairing, "load-key");
  await bootStep(reloadLinks, "refresh-links");
  // Read once here too: refreshInbox only runs paired, and the Slide picker
  // is there whether or not a key has been pasted yet.
  await bootStep(
    () =>
      refreshSlideOptions(insertSlideSelect, (error) => {
        details.addFailure(error, "The slide list was not refreshed");
      }),
    "refresh-slide",
  );
  // Unpaired is not a boot failure: the Inbox says so itself, and the Links
  // list works without a key.
  if (workspace !== null) await bootStep(refreshInbox, "refresh-inbox");
});
