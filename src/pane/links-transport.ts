// src/pane/links-transport.ts
// The Excel side of local mode: which way this device's links travel (the
// per-device setting, src/link/transport-setting.ts) and the synchronous
// clipboard copy an export or a copy runs through when they travel locally.
// Office.js never reaches here - the caller's run() does the actual Excel
// work; this module only starts the clipboard write, waits on it and falls
// back to a manual copy. Invariant: copy() starts the clipboard write before
// its own first await (WebKit only allows the write inside the click's own
// call stack).

import { BUNDLE_MAX_CHARS, encodeBundle } from "../link/bundle";
import { localWorkspace } from "../link/local";
import { LocalCollector } from "../link/local-collector";
import type { PushSummary } from "../excel";
import type { RelayApi } from "../link/relay";
import {
  loadTransport,
  saveTransport,
  type LinkTransport,
} from "../link/transport-setting";
import type { KeyStore, Workspace } from "../link/workspace";
import { beginClipboardWrite, copyBundleNow } from "../ui/clipboard-links";
import type { Toast } from "../ui/toast";

const TOO_MUCH_MESSAGE =
  "That is too much to copy at once: copy one project or the selected links.";
const NOTHING_MESSAGE = "Nothing to copy: no link could be read.";
const NOTHING_PREPARED_MESSAGE = "Nothing is waiting to be copied.";
const COPY_FAILED_MESSAGE =
  "Copy failed: select the text below and copy it by hand (Ctrl+C, or ⌘C on a Mac).";
const COPIED_MESSAGE =
  "Copied for PowerPoint. Paste it in the pls,fix pane, Inbox tab.";

// A press that lands before boot() has read the setting meets one sentence,
// never a crash. boot() replaces the placeholder long before a modeller can
// reach a button, so no test pins the string down.
const TRANSPORT_LOADING_MESSAGE = "Still starting up: try again in a moment.";

export interface LinksTransportDeps {
  root: ParentNode;
  keyStore: KeyStore;
  toast: Toast;
}

export interface LinksTransport {
  mode(): LinkTransport;
  setMode(mode: LinkTransport): Promise<void>;
  // Local mode only. Must be called synchronously inside the press: it starts
  // the clipboard write before its first await, then runs the Excel work
  // against a LocalCollector and the local workspace.
  copy<T>(
    run: (ws: Workspace, relay: RelayApi) => Promise<T>,
    lines: { copied: (result: T) => string; ready: (result: T) => string },
  ): Promise<string>;
  // The "Copy for PowerPoint" button: retries the last prepared copy through
  // the synchronous copy-event fallback. Throws when nothing is prepared.
  retryCopy(): string;
}

interface Elements {
  bar: HTMLElement;
  text: HTMLElement;
  manual: HTMLTextAreaElement;
  select: HTMLSelectElement;
  linkKeySection: HTMLElement;
  autopushRow: HTMLElement;
  pushSelected: HTMLElement;
  pushAll: HTMLElement;
  localHint: HTMLElement;
  relayHint: HTMLElement;
}

interface State {
  deps: LinksTransportDeps;
  els: Elements;
  mode: LinkTransport;
  // The last bundle a copy() prepared but could not land on the clipboard by
  // itself: what "Copy for PowerPoint" retries. Null once copied, or once
  // nothing has been prepared yet.
  prepared: string | null;
}

export async function installLinksTransport(
  deps: LinksTransportDeps,
): Promise<LinksTransport> {
  const state: State = {
    deps,
    els: elements(deps.root),
    mode: await loadTransport(deps.keyStore),
    prepared: null,
  };
  // Before the tab's first render: a boot that reads relay must never flash
  // local's sections first.
  applyMode(state, state.mode);
  return {
    mode: () => state.mode,
    setMode: (mode) => setMode(state, mode),
    copy: (run, lines) => copy(state, run, lines),
    retryCopy: () => retryCopy(state),
  };
}

async function setMode(state: State, mode: LinkTransport): Promise<void> {
  await saveTransport(state.deps.keyStore, mode);
  state.mode = mode;
  applyMode(state, mode);
}

// Visibility and labels only: local mode's own auto-push switch-off is
// src/pane/links-tab.ts's job (it alone holds the real relay and the
// toggle), run beside this from both the boot order and the select's
// change handler.
function applyMode(state: State, mode: LinkTransport): void {
  const local = mode === "local";
  state.els.select.value = mode;
  state.els.linkKeySection.hidden = local;
  state.els.autopushRow.hidden = local;
  state.els.pushSelected.textContent = local
    ? "Copy selected"
    : "Push selected";
  state.els.pushAll.textContent = local ? "Copy all" : "Push all";
  state.els.localHint.hidden = !local;
  state.els.relayHint.hidden = local;
  if (!local) hideBar(state);
}

async function copy<T>(
  state: State,
  run: (ws: Workspace, relay: RelayApi) => Promise<T>,
  lines: { copied: (result: T) => string; ready: (result: T) => string },
): Promise<string> {
  const pending = beginClipboardWrite();
  const collector = new LocalCollector();
  let result: T;
  try {
    result = await run(await localWorkspace(), collector);
  } catch (error) {
    pending.reject(error);
    throw error;
  }

  const json = encodeBundle(collector.bundle());
  if (json.length > BUNDLE_MAX_CHARS) {
    const error = new Error(TOO_MUCH_MESSAGE);
    pending.reject(error);
    throw error;
  }
  if (collector.isEmpty()) {
    pending.reject(new Error(NOTHING_MESSAGE));
    return NOTHING_MESSAGE;
  }

  pending.resolve(json);
  if (await pending.done) {
    hideBar(state);
    return lines.copied(result);
  }
  showBar(state, json, lines.ready(result));
  return lines.ready(result);
}

function retryCopy(state: State): string {
  const json = state.prepared;
  if (json === null) throw new Error(NOTHING_PREPARED_MESSAGE);
  if (copyBundleNow(json)) {
    hideBar(state);
    return COPIED_MESSAGE;
  }
  showManualBox(state, json);
  throw new Error(COPY_FAILED_MESSAGE);
}

function showBar(state: State, json: string, text: string): void {
  state.prepared = json;
  state.els.bar.hidden = false;
  state.els.text.textContent = text;
  state.els.manual.hidden = true;
}

function hideBar(state: State): void {
  state.prepared = null;
  state.els.bar.hidden = true;
  state.els.manual.hidden = true;
  state.els.manual.value = "";
}

function showManualBox(state: State, json: string): void {
  state.els.manual.value = json;
  state.els.manual.hidden = false;
  state.els.manual.select();
}

function elements(root: ParentNode): Elements {
  return {
    bar: element(root, "copy-ready"),
    text: element(root, "copy-ready-text"),
    manual: element(root, "copy-manual"),
    select: element(root, "link-transport"),
    linkKeySection: element(root, "link-key-section"),
    autopushRow: element(root, "links-autopush-row"),
    pushSelected: element(root, "push-selected"),
    pushAll: element(root, "push-all"),
    localHint: element(root, "transport-local-hint"),
    relayHint: element(root, "transport-relay-hint"),
  };
}

function element<T extends Element>(root: ParentNode, id: string): T {
  const found = root.querySelector<T>(`#${id}`);
  if (!found) throw new Error(`Missing element #${id}`);
  return found;
}

// What the Links tab holds until boot() has read the setting: every call
// answers TRANSPORT_LOADING_MESSAGE, so no call site has to null-check.
export function pendingLinksTransport(): LinksTransport {
  const notReady = (): never => {
    throw new Error(TRANSPORT_LOADING_MESSAGE);
  };
  return {
    mode: notReady,
    setMode: () => Promise.reject(new Error(TRANSPORT_LOADING_MESSAGE)),
    copy: () => Promise.reject(new Error(TRANSPORT_LOADING_MESSAGE)),
    retryCopy: notReady,
  };
}

// An export's two lines (a range, a table, a text link or a chart). A chart
// Excel could not describe keeps the note relay mode shows beside its label.
export function exportCopyLines(): {
  copied: (result: { label: string; note?: string }) => string;
  ready: (result: { label: string; note?: string }) => string;
} {
  const named = (result: { label: string; note?: string }): string =>
    result.note === undefined
      ? result.label
      : `${result.label} (${result.note})`;
  return {
    copied: (result) =>
      `Copied for PowerPoint: ${named(result)}. Paste it in PowerPoint: pls,fix, Inbox tab.`,
    ready: (result) =>
      `${named(result)} is linked and ready: press Copy for PowerPoint.`,
  };
}

// Copy selected / Copy all: the push counts, said as copies.
export function pushCopyLines(): {
  copied: (summary: PushSummary) => string;
  ready: (summary: PushSummary) => string;
} {
  const counts = (summary: PushSummary): string =>
    `${summary.pushed} copied, ${summary.missing} missing, ${summary.failed} failed`;
  return {
    copied: (summary) =>
      `${counts(summary)}. Paste it in PowerPoint: pls,fix, Inbox tab.`,
    ready: (summary) => `${counts(summary)}: press Copy for PowerPoint.`,
  };
}
