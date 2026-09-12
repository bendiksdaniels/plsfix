// The Brand tab's "Keyboard shortcuts" section: one row per remappable action
// read off the served shortcuts.json, an Apply that writes the user's own key
// map and a Reset all that takes every customisation back. Owns every
// Office.actions call the add-in makes; the grammar and the maps it sends are
// src/shortcuts-model.ts. Invariant: nothing calls Office.actions until
// isSetSupported("KeyboardShortcuts", "1.1") has said yes.

import {
  diffShortcuts,
  duplicateKeys,
  parseCombo,
  parseShortcutRows,
  showKey,
  type KeyPlatform,
  type ShortcutRow,
} from "../shortcuts-model";

// Same origin, served straight out of public/ beside the pane in every
// deployment, so the printable card and this panel read one file.
const SHORTCUTS_URL = "shortcuts.json";
// A webview that never answers must not leave the section empty for ever, nor
// the pane busy: both round trips get a deadline of their own.
const FETCH_TIMEOUT_MS = 5_000;
const API_TIMEOUT_MS = 15_000;

const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";
const LIST_FAILED =
  "The shortcut list could not be read. Reopen the pane to try again.";
const HINT =
  "Type a combination beside any action, then Apply. An empty box keeps the pls,fix default.";
const APPLIED = "Shortcuts updated for your account.";
const RESET_DONE = "Shortcuts are back to the pls,fix defaults.";
const SLOW = "Office did not answer. Try again in a moment.";

let rows: ShortcutRow[] = [];
let platform: KeyPlatform = "windows";

/**
 * Renders the section, wires the printable-card button to the pane's own
 * dialog opener and reads the user's keys back whenever the Brand tab opens.
 * A pane without the section (the deck's) is left alone.
 */
export async function installShortcutsPanel(
  root: ParentNode,
  onCard: () => void,
): Promise<void> {
  const list = root.querySelector<HTMLElement>("#shortcuts-list");
  if (list === null) return;
  platform = platformOf();
  root
    .querySelector<HTMLButtonElement>("#shortcuts-card")
    ?.addEventListener("click", onCard);
  // The section is installed before Office is ready, so the requirement set
  // is asked again every time the tab that shows it comes into view.
  root
    .querySelector<HTMLElement>("#tab-brand")
    ?.addEventListener("click", () => void refresh(root));

  try {
    rows = parseShortcutRows(await fetchShortcuts());
  } catch {
    rows = [];
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
  const custom = boxes();
  refuseBadBox(custom);

  const map = diffShortcuts(rows, custom);
  const repeated = duplicateKeys(map);
  if (repeated.length > 0) {
    throw new Error(`${repeated[0]!} is set on two actions.`);
  }

  const wanted = Object.values(map).filter((key) => key !== null);
  const clashes = await clashesAmong(api, wanted);
  await send(api, map);
  await fill(api);
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
  await send(api, map);
  for (const row of rows) {
    const input = box(row.id);
    if (input !== null) input.value = "";
  }
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

// Office.PlatformType is a string enum ("Mac"), so the value survives a host
// that has not defined the enum object yet.
function platformOf(): KeyPlatform {
  try {
    if (String(Office.context.platform) === "Mac") return "mac";
  } catch {
    // No host: the browser's own answer is the best there is.
  }
  return /Mac/i.test(navigator.userAgent) ? "mac" : "windows";
}

async function refresh(root: ParentNode): Promise<void> {
  if (rows.length === 0) return;
  const api = actions();
  say(root, api === null ? UNSUPPORTED : HINT);
  for (const row of rows) {
    const input = box(row.id);
    if (input !== null) input.disabled = api === null;
  }
  if (api !== null) await fill(api);
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

function rowLabel(row: ShortcutRow): HTMLLabelElement {
  const shown = showKey(row.defaultKey, platform);
  const label = document.createElement("label");
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

// getShortcuts answers the manifest's own key where nothing was customised, so
// a box shows a value only when it genuinely differs from the shipped one.
async function fill(api: Office.Actions): Promise<void> {
  let current: Record<string, string | null>;
  try {
    current = await deadline(api.getShortcuts());
  } catch {
    return;
  }
  for (const row of rows) {
    const input = box(row.id);
    if (input === null) continue;
    const key = current[row.id];
    const custom =
      typeof key === "string" && !sameKey(key, row.defaultKey) ? key : "";
    input.value = custom === "" ? "" : showKey(custom, platform);
  }
}

function sameKey(one: string, other: string): boolean {
  return showKey(one, "windows") === showKey(other, "windows");
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
    // Nothing is left to blame on the map: every id came from shortcuts.json
    // and every value parsed. Microsoft documents one environmental refusal
    // for this call, an anonymous user, and names no error code for it.
    throw new Error(UNSUPPORTED, { cause: error });
  }
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
      reject(new Error(SLOW));
    }, API_TIMEOUT_MS);
    work.then(resolve, reject).finally(() => {
      window.clearTimeout(timer);
    });
  });
}
