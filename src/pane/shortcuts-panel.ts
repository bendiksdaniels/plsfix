// The Brand tab's "Keyboard shortcuts" section: a box per action, Apply, Reset
// all and the printable card. Owns every Office.actions call the add-in makes;
// the grammar and the maps are src/shortcuts-model.ts. Invariant: nothing is
// written until the panel has read back what Office currently holds.

import {
  diffShortcuts,
  firstClash,
  parseCombo,
  parseShortcutRows,
  showKey,
  type ShortcutRow,
} from "../shortcuts-model";

// Same origin, served straight out of public/ beside the pane in every
// deployment, so the printable card and this panel read one file.
const SHORTCUTS_URL = "shortcuts.json";
// A webview that never answers must not leave the section empty for ever, nor
// the pane busy: both round trips get a deadline of their own.
const FETCH_TIMEOUT_MS = 5_000;
const API_TIMEOUT_MS = 15_000;
// Marks this module's own deadline apart from anything Office rejects with.
const TIMED_OUT = "plsfixShortcutTimeout";

const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";
const LIST_FAILED =
  "The shortcut list could not be read. Reopen the pane to try again.";
const READ_FAILED =
  "Your current shortcuts could not be read, so nothing was changed. Reopen the Brand tab and try again.";
const HINT =
  "Type a combination beside any action, then Apply. An empty box keeps the pls,fix default.";
const APPLIED = "Shortcuts updated for your account.";
const RESET_DONE = "Shortcuts are back to the pls,fix defaults.";
const SLOW = "Office did not answer. Try again in a moment.";
const SEND_SLOW =
  "Office did not answer. Your keys may still have been saved: reopen the Brand tab to check.";

let rows: ShortcutRow[] = [];
// True once a getShortcuts read has landed: until then the boxes are this
// module's own blanks, not the user's intent, and may not be written back.
let filled = false;
// What this module last wrote into each box, so a box the user has typed into
// since is left alone by the next read.
const lastFilled = new Map<string, string>();

/**
 * Renders the section, wires the printable-card button to the pane's own
 * dialog opener and re-reads the user's keys whenever the Brand tab becomes
 * visible. A pane without the section (the deck's) is left alone.
 */
export async function installShortcutsPanel(
  root: ParentNode,
  onCard: () => void,
): Promise<void> {
  const list = root.querySelector<HTMLElement>("#shortcuts-list");
  if (list === null) return;
  rows = [];
  filled = false;
  lastFilled.clear();
  root
    .querySelector<HTMLButtonElement>("#shortcuts-card")
    ?.addEventListener("click", onCard);
  watchTab(root, list);

  try {
    rows = parseShortcutRows(await fetchShortcuts());
  } catch {
    say(root, LIST_FAILED);
    return;
  }
  list.replaceChildren(...rows.map(rowLabel));
  await refresh(root);
}

/** Writes what the boxes hold, after Office has been asked about clashes. */
export async function applyShortcuts(): Promise<string> {
  const api = requireApi();
  requireRows();
  // The boxes are the user's intent only once they hold what Office has: an
  // Apply over a panel that never read would send 42 nulls and wipe the lot.
  if (!filled && !(await fill(api, false))) throw new Error(READ_FAILED);

  const custom = boxes();
  refuseBadBox(custom);
  const map = diffShortcuts(rows, custom);
  const clash = firstClash(rows, map);
  if (clash !== null) {
    throw new Error(
      `${clash.first} and ${clash.second} would both use ${clash.key}.`,
    );
  }

  const wanted = Object.values(map).filter((key) => key !== null);
  const clashes = await clashesAmong(api, wanted);
  await send(api, map);
  await fill(api, true);
  return clashes.length === 0
    ? APPLIED
    : `${APPLIED} Already used elsewhere: ${clashes.join(", ")}.`;
}

/** Takes every customisation back: null for every action Office knows. */
export async function resetShortcuts(): Promise<string> {
  const api = requireApi();
  requireRows();
  const map: Record<string, string | null> = {};
  for (const row of rows) map[row.id] = null;
  // No read first: reverting everything is the one intent that does not depend
  // on what the user currently has.
  await send(api, map);
  for (const row of rows) {
    const input = box(row.id);
    if (input !== null) input.value = "";
    lastFilled.set(row.id, "");
  }
  filled = true;
  return RESET_DONE;
}

// The requirement set, asked through a try/catch: a plain browser has no
// Office at all and a host mid-boot can throw rather than answer false.
function actions(): Office.Actions | null {
  try {
    if (typeof Office === "undefined") return null;
    if (!Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1"))
      return null;
    return Office.actions;
  } catch {
    return null;
  }
}

function requireApi(): Office.Actions {
  const api = actions();
  if (api === null) throw new Error(UNSUPPORTED);
  return api;
}

function requireRows(): void {
  if (rows.length === 0) throw new Error(LIST_FAILED);
}

// src/ui/tabs.ts shows a panel by clearing its `hidden`, whether the tab was
// clicked, arrowed to, or activated by the tool search - so the panel watches
// for that, never for a click on the tab button.
function watchTab(root: ParentNode, list: HTMLElement): void {
  const panel = list.closest<HTMLElement>('[role="tabpanel"]');
  if (panel === null || typeof MutationObserver === "undefined") return;
  new MutationObserver(() => {
    if (!panel.hidden) void refresh(root);
  }).observe(panel, { attributes: true, attributeFilter: ["hidden"] });
}

async function refresh(root: ParentNode): Promise<void> {
  if (rows.length === 0) return;
  const api = actions();
  say(root, api === null ? UNSUPPORTED : HINT);
  for (const row of rows) {
    const input = box(row.id);
    if (input !== null) input.disabled = api === null;
  }
  if (api !== null) await fill(api, false);
}

function say(root: ParentNode, sentence: string): void {
  const note = root.querySelector<HTMLElement>("#shortcuts-note");
  if (note !== null) note.textContent = sentence;
}

function inputId(id: string): string {
  return `shortcut-${id}`;
}

function box(id: string): HTMLInputElement | null {
  const input = document.getElementById(inputId(id));
  return input instanceof HTMLInputElement ? input : null;
}

// One full-width row per action: the name and its shipped key on the first
// line, the box under it. Two columns would clip - a kbd does not wrap and the
// longest key is wider than half a 320 px dock.
function rowLabel(row: ShortcutRow): HTMLLabelElement {
  const shown = showKey(row.defaultKey);
  const label = document.createElement("label");
  label.className = "key-label";
  const name = document.createElement("span");
  name.textContent = `${row.name} `;
  const key = document.createElement("kbd");
  key.textContent = shown;
  name.append(key);

  const input = document.createElement("input");
  input.type = "text";
  input.className = "key-input";
  input.id = inputId(row.id);
  input.placeholder = shown;
  input.autocomplete = "off";
  input.spellcheck = false;
  input.setAttribute("aria-label", `Custom key for ${row.name}`);

  label.append(name, input);
  return label;
}

function boxes(): Record<string, string> {
  const custom: Record<string, string> = {};
  for (const row of rows) {
    const input = box(row.id);
    if (input !== null) custom[row.id] = input.value;
  }
  return custom;
}

// Named, focused and refused one box at a time: the next Apply names the next
// one, which beats a toast listing four rows nobody can read at 320 px.
function refuseBadBox(custom: Record<string, string>): void {
  for (const row of rows) {
    const raw = (custom[row.id] ?? "").trim();
    if (raw === "") continue;
    const parsed = parseCombo(raw);
    if (parsed.ok) continue;
    box(row.id)?.focus();
    throw new Error(`${row.name}: ${parsed.reason}`);
  }
}

/**
 * Reads the account's keys into the boxes. getShortcuts answers the manifest's
 * own key where nothing was customised, so a box shows a value only where it
 * genuinely differs. Answers whether the read landed.
 */
async function fill(api: Office.Actions, force: boolean): Promise<boolean> {
  let current: Record<string, string | null>;
  try {
    current = await deadline(api.getShortcuts());
  } catch {
    // A read that fails after an earlier success drops the latch too, so the
    // next Apply reads again instead of trusting a map that may be stale.
    filled = false;
    return false;
  }
  for (const row of rows) {
    const input = box(row.id);
    if (input === null) continue;
    // A box typed into since the last read is unapplied intent: only Apply's
    // own re-read, which knows the write landed, may overwrite it.
    if (!force && input.value !== (lastFilled.get(row.id) ?? "")) continue;
    const key = current[row.id];
    const shown =
      typeof key === "string" && !sameKey(key, row.defaultKey)
        ? showKey(key)
        : "";
    input.value = shown;
    lastFilled.set(row.id, shown);
  }
  filled = true;
  return true;
}

function sameKey(one: string, other: string): boolean {
  return showKey(one) === showKey(other);
}

async function clashesAmong(
  api: Office.Actions,
  wanted: string[],
): Promise<string[]> {
  if (wanted.length === 0) return [];
  try {
    const report = await deadline(api.areShortcutsInUse(wanted));
    return report.filter((item) => item.inUse).map((item) => item.shortcut);
  } catch {
    // A host that will not answer the clash question must not block the
    // remap: Office asks the user itself the first time a clash is pressed.
    return [];
  }
}

async function send(
  api: Office.Actions,
  map: Record<string, string | null>,
): Promise<void> {
  try {
    // The typings declare the value as string; both documentation pages
    // document null as "take the customisation back, keep the shipped key",
    // and null is the only way this panel can undo one. Cast once, here.
    await deadline(api.replaceShortcuts(map as Record<string, string>));
  } catch (error) {
    // A swallowed batch is not a refusal: the write may still land, so say so
    // rather than blaming an account that is signed in perfectly well.
    if (isTimeout(error)) throw carry(SEND_SLOW, error);
    // Nothing is left to blame on the map: every id came from shortcuts.json
    // and every value parsed. Microsoft documents one environmental refusal
    // for this call, an anonymous user, and names no error code for it.
    throw carry(UNSUPPORTED, error);
  }
}

// src/ui/report.ts reads `code` and `debugInfo` off the error it is handed and
// never walks `.cause`, so Office's own diagnostics travel on this one too.
function carry(message: string, cause: unknown): Error {
  const error = new Error(message, { cause });
  for (const key of ["code", "debugInfo"]) {
    const value = readProp(cause, key);
    if (value !== undefined) Object.assign(error, { [key]: value });
  }
  return error;
}

function readProp(value: unknown, key: string): unknown {
  if (typeof value !== "object" || value === null || !(key in value))
    return undefined;
  return (value as Record<string, unknown>)[key];
}

async function fetchShortcuts(): Promise<unknown> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => {
    controller.abort();
  }, FETCH_TIMEOUT_MS);
  try {
    const url = new URL(SHORTCUTS_URL, location.href).href;
    const response = await fetch(url, { signal: controller.signal });
    if (!response.ok) {
      throw new Error(`fetch shortcuts.json: HTTP ${String(response.status)}`);
    }
    return await response.json();
  } finally {
    window.clearTimeout(timer);
  }
}

function deadline<T>(work: Promise<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(() => {
      reject(Object.assign(new Error(SLOW), { [TIMED_OUT]: true }));
    }, API_TIMEOUT_MS);
    work.then(resolve, reject).finally(() => {
      window.clearTimeout(timer);
    });
  });
}

function isTimeout(error: unknown): boolean {
  return readProp(error, TIMED_OUT) === true;
}
