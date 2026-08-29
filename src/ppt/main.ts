// PowerPoint pane: boot (Office.js, the PowerPoint and PowerPointApi 1.5
// gates, toast, error reporting, tabs, version), the pane's state - the links
// this deck holds, which are ticked, and the paired workspace - and one guarded
// handler per button. Drawing is views.ts, selection maths actions.ts, and
// every deck or relay call goes through links.ts / host.ts.

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
import { makeGuard } from "../ui/guard";
import { describeError, installErrorReporting } from "../ui/report";
import { installTabs } from "../ui/tabs";
import { createToast } from "../ui/toast";
import { formatVersion } from "../ui/version";
import {
  pruneSelection,
  requireSelection,
  selectedRows,
  slideRows,
  toRowViews,
  updateDetails,
} from "./actions";
import { activeSlideId, breakLink, goToSlide } from "./host";
import {
  insertFromInbox,
  listInbox,
  listLinks,
  summarize,
  updateLinks,
  type LinkRow,
} from "./links";
import { revertLinks, summarizeRevert } from "./revert";
import { renderInbox, renderLinkRows } from "./views";

const APP_VERSION = formatVersion(__APP_VERSION__);
const REPORT_CONTEXT = { host: "PowerPoint", version: APP_VERSION };
const NOT_PAIRED =
  "Not paired: paste the link key from Excel > Links > Settings";
const PAIR_FIRST = "Paste the link key in Settings";
const NOT_CONNECTED = "PowerPoint is not connected.";

const getElement = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as T;
};

const connectionStatus = getElement<HTMLSpanElement>("connection-status");
const linkRowsBody = getElement<HTMLTableSectionElement>("link-rows");
const linksEmpty = getElement("links-empty");
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
installTabs(getElement("tab-bar"));

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

function renderLinks(): void {
  renderLinkRows(linkRowsBody, toRowViews(rows, selected), toggleSelection);
  linksEmpty.hidden = rows.length > 0;
}

function toggleSelection(key: string, isSelected: boolean): void {
  if (isSelected) selected.add(key);
  else selected.delete(key);
}

// The list is hidden rather than emptied when unpaired, and its children are
// replaced either way: a stale Insert button must never survive a re-render.
function renderInboxView(): void {
  const paired = workspace !== null;
  renderInbox(inboxList, inboxItems, (item) => {
    act(() => insertItem(item), "insert-link");
  });
  inboxList.hidden = !paired;
  inboxUnpaired.hidden = paired;
}

function renderPairing(): void {
  workspaceState.textContent = workspace === null ? NOT_PAIRED : "Paired";
}

// ---------------------------------------------------------------------------
// The guard every button runs through
// ---------------------------------------------------------------------------

// makeGuard's success path notifies with a message only, so an action with
// per-link lines to show stages them here; the guard clears them afterwards.
let stagedDetails: string | undefined;

function noteDetail(line: string): void {
  stagedDetails =
    stagedDetails === undefined ? line : `${stagedDetails}\n${line}`;
}

function setBusy(busy: boolean): void {
  const buttons =
    document.querySelectorAll<HTMLButtonElement>(".app-shell button");
  for (const button of buttons) button.disabled = busy;
}

const guard = makeGuard({
  setBusy,
  notify: (message, kind, details) =>
    toast.show(message, kind, details ?? stagedDetails),
  describe: (error, action) => describeError(error, REPORT_CONTEXT, action),
  finally: () => {
    stagedDetails = undefined;
  },
});

// Set once Office.onReady has confirmed PowerPoint and PowerPointApi 1.5.
let ready = false;

// Every button is live from the first paint, so a click during a cold boot - or
// in a deck on a PowerPoint the pane rejected - would reach PowerPoint.run and
// toast a raw "Cannot read properties of undefined". Every handler starts here
// instead, and the connection badge is not the only thing that says so.
function act(run: () => Promise<string>, action: string): void {
  void guard(async () => {
    if (!ready) throw new Error(NOT_CONNECTED);
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
// leaves the list as it was and says so in the details rather than replacing
// the report with its own error.
async function refreshQuietly(): Promise<void> {
  try {
    await reloadLinks();
  } catch (error) {
    noteDetail(
      `The list was not refreshed: ${describeError(error, REPORT_CONTEXT).message}`,
    );
  }
}

async function refreshLinks(): Promise<string> {
  await reloadLinks();
  if (rows.length === 0) return "No linked objects in this deck.";
  return `${String(rows.length)} linked ${rows.length === 1 ? "object" : "objects"}.`;
}

async function updateRows(subset: LinkRow[]): Promise<string> {
  const summary = await updateLinks(subset, relay);
  stagedDetails = updateDetails(summary);
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
  stagedDetails =
    summary.failures.length > 0 ? summary.failures.join("\n") : undefined;
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
  inboxItems = await listInbox(ws, relay);
  renderInboxView();
  if (inboxItems.length === 0) return "Nothing waiting from Excel.";
  return `${String(inboxItems.length)} waiting to insert.`;
}

async function insertItem(item: InboxItem): Promise<string> {
  const ws = requireWorkspace();
  await insertFromInbox(item, ws, relay);
  // The relay copy is gone, so the item leaves the list without a second call.
  inboxItems = inboxItems.filter((waiting) => waiting.id !== item.id);
  renderInboxView();
  await refreshQuietly();
  return `Inserted ${item.label}.`;
}

async function saveKey(): Promise<string> {
  const key = workspaceKey.value.trim();
  if (key === "") throw new Error("Paste the link key from Excel first.");
  workspace = await pairWith(key);
  // The key is the secret itself: it is stored, never echoed back.
  workspaceKey.value = "";
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
    noteDetail(
      `The inbox was not read: ${describeError(error, REPORT_CONTEXT).message}`,
    );
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
  "save-key": saveKey,
  "forget-key": forgetKey,
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

// Unpaired is not a boot failure: the Inbox says so itself, and the Links list
// works without a key.
async function loadInbox(): Promise<void> {
  if (workspace === null) return;
  await refreshInbox();
}

Office.onReady(async ({ host }) => {
  if (host !== Office.HostType.PowerPoint) {
    connectionStatus.textContent = "PowerPoint required";
    connectionStatus.className = "connection error";
    return;
  }

  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    connectionStatus.textContent = "PowerPoint 2021 / Microsoft 365 required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "PowerPoint connected";
  connectionStatus.className = "connection ready";
  ready = true;

  await bootStep(loadPairing, "load-key");
  await bootStep(reloadLinks, "refresh-links");
  await bootStep(loadInbox, "refresh-inbox");
});
