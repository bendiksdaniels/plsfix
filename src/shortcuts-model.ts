// The pure model behind the shortcut manager: public/shortcuts.json read into
// printable rows, the key grammar Microsoft documents, and the map
// Office.actions.replaceShortcuts takes. Invariant: a combination is printed
// exactly as Office takes it, so no modifier is ever relabelled per platform.

/** One remappable action: what shortcuts.json calls it and its shipped key. */
export interface ShortcutRow {
  id: string;
  name: string;
  defaultKey: string;
}

/**
 * The modifier names Office's key strings use. Cmd is its own modifier, not an
 * alias: the docs say it is supported on macOS and mapped to Ctrl on Windows,
 * so folding it away would stop a Mac user setting a Command shortcut at all.
 */
export type Modifier = "Ctrl" | "Cmd" | "Shift" | "Alt";

/** A validated key combination: at least one modifier plus exactly one key. */
export interface Combo {
  modifiers: Modifier[];
  key: string;
}

/** Either the combination, or the one sentence the pane shows instead. */
export type ComboResult =
  { ok: true; combo: Combo } | { ok: false; reason: string };

/** Two actions that would answer to one key once a map is applied. */
export interface KeyClash {
  key: string;
  first: string;
  second: string;
}

/**
 * Every refusal says the same thing: the grammar is short enough that naming
 * which of its rules was broken would be longer than the rule itself.
 */
export const INVALID_COMBO_REASON =
  "Use Ctrl, Alt or Shift plus one key, for example Ctrl+Shift+A.";

// Office writes Ctrl+Shift+Alt+x in its own good-practice pattern, and
// public/shortcuts.json follows it, so that is the order this module prints.
const MODIFIER_ORDER: Modifier[] = ["Ctrl", "Cmd", "Shift", "Alt"];

// The names a typist may reach for. Option is the Mac name of Alt ("On macOS,
// the Alt key is mapped to the Option key") and folds onto it; Command is only
// the long spelling of Cmd, which stays a modifier of its own.
const MODIFIER_NAMES = new Map<string, Modifier>([
  ["ctrl", "Ctrl"],
  ["control", "Ctrl"],
  ["cmd", "Cmd"],
  ["command", "Cmd"],
  ["shift", "Shift"],
  ["alt", "Alt"],
  ["option", "Alt"],
  ["opt", "Alt"],
]);

// "Key combinations can include characters A-Z, a-z, 0-9, and the punctuation
// marks -, _ and +" - and "-" and "_" sit on one physical key, so they are the
// same shortcut and fold to one.
const KEY_CHARACTER = /^[A-Za-z0-9\-_+]$/;

// Not in that character list, but Microsoft's own manifest samples bind
// Ctrl+Alt+Up and Ctrl+Alt+Down, so the four arrows are keys too.
const ARROW_KEYS = new Map<string, string>([
  ["up", "Up"],
  ["down", "Down"],
  ["left", "Left"],
  ["right", "Right"],
]);

interface RawAction {
  id?: unknown;
  name?: unknown;
}

interface RawBinding {
  action?: unknown;
  key?: { default?: unknown } | unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function defaultKeyOf(binding: RawBinding): string | null {
  const key: unknown = binding.key;
  if (!isRecord(key) || typeof key.default !== "string") return null;
  return key.default;
}

/**
 * The remappable rows in the order shortcuts.json binds them, which is the
 * order the printable card prints. An action nobody bound a key to has no
 * default to show and is left out.
 */
export function parseShortcutRows(source: unknown): ShortcutRow[] {
  if (
    !isRecord(source) ||
    !Array.isArray(source.actions) ||
    !Array.isArray(source.shortcuts)
  ) {
    throw new Error("parse shortcuts.json: no actions and shortcuts arrays");
  }

  const names = new Map<string, string>();
  for (const action of source.actions as RawAction[]) {
    if (typeof action.id === "string" && typeof action.name === "string") {
      names.set(action.id, action.name);
    }
  }

  const rows: ShortcutRow[] = [];
  for (const binding of source.shortcuts as RawBinding[]) {
    const id: unknown = binding.action;
    const defaultKey = defaultKeyOf(binding);
    if (typeof id !== "string" || defaultKey === null) continue;
    const name = names.get(id);
    if (name === undefined) {
      throw new Error(`parse shortcuts.json: no action named ${id}`);
    }
    rows.push({ id, name, defaultKey });
  }
  return rows;
}

/** "Ctrl+Shift+A" -> the validated combination, else the one reason sentence. */
export function parseCombo(text: string): ComboResult {
  const parts = splitCombo(text);
  if (parts === null) return { ok: false, reason: INVALID_COMBO_REASON };

  const modifiers: Modifier[] = [];
  let key: string | null = null;
  for (const part of parts) {
    const modifier = MODIFIER_NAMES.get(part.toLowerCase());
    if (modifier !== undefined) {
      if (modifiers.includes(modifier))
        return { ok: false, reason: INVALID_COMBO_REASON };
      modifiers.push(modifier);
      continue;
    }
    const named = keyOf(part);
    if (key !== null || named === null)
      return { ok: false, reason: INVALID_COMBO_REASON };
    key = named;
  }

  // At least one modifier, exactly one key, and Shift never on its own.
  if (key === null || modifiers.length === 0)
    return { ok: false, reason: INVALID_COMBO_REASON };
  if (modifiers.length === 1 && modifiers[0] === "Shift")
    return { ok: false, reason: INVALID_COMBO_REASON };

  return {
    ok: true,
    combo: {
      modifiers: MODIFIER_ORDER.filter((name) => modifiers.includes(name)),
      key,
    },
  };
}

function keyOf(part: string): string | null {
  const arrow = ARROW_KEYS.get(part.toLowerCase());
  if (arrow !== undefined) return arrow;
  if (!KEY_CHARACTER.test(part)) return null;
  return part === "_" ? "-" : part.toUpperCase();
}

// "+" is both the separator and a legal key, so a combination ending in two of
// them (spaced or not) has "+" for its key; every other empty part is a typo.
function splitCombo(text: string): string[] | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  const plusKey = /^(.*?)\s*\+\s*\+\s*$/.exec(trimmed);
  const body = plusKey ? plusKey[1]! : trimmed;
  const parts = body.split("+").map((part) => part.trim());
  if (parts.some((part) => part === "")) return null;
  return plusKey ? [...parts, "+"] : parts;
}

/** The one printed form, which is also the one Office is handed. */
export function formatCombo(combo: Combo): string {
  return [...combo.modifiers, combo.key].join("+");
}

/**
 * A stored key string as the pane prints it. One Office writes in a form this
 * grammar does not cover (a page key, say) is shown exactly as it stands rather
 * than swallowed.
 */
export function showKey(key: string): string {
  const parsed = parseCombo(key);
  return parsed.ok ? formatCombo(parsed.combo) : key;
}

/**
 * The boxes the pane holds turned into replaceShortcuts' map: the canonical
 * combination where the user set one, and null - back to the shipped default -
 * for a box left empty or retyped as the default it already had.
 */
export function diffShortcuts(
  defaults: ShortcutRow[],
  custom: Record<string, string>,
): Record<string, string | null> {
  const map: Record<string, string | null> = {};
  for (const row of defaults) {
    const raw = (custom[row.id] ?? "").trim();
    if (raw === "") {
      map[row.id] = null;
      continue;
    }
    const parsed = parseCombo(raw);
    if (!parsed.ok) {
      throw new Error(`shortcut ${row.id}: ${parsed.reason}`);
    }
    const wanted = formatCombo(parsed.combo);
    map[row.id] = wanted === showKey(row.defaultKey) ? null : wanted;
  }
  return map;
}

/**
 * The first pair of actions a map would leave on one key, comparing what each
 * action would EFFECTIVELY answer to: its custom key where the map sets one,
 * its shipped default otherwise. A custom key landing on another action's
 * untouched default is the clash neither Office nor the user can see coming.
 */
export function firstClash(
  defaults: ShortcutRow[],
  map: Record<string, string | null>,
): KeyClash | null {
  const owner = new Map<string, string>();
  for (const row of defaults) {
    const key = map[row.id] ?? showKey(row.defaultKey);
    const first = owner.get(key);
    if (first !== undefined) return { key, first, second: row.name };
    owner.set(key, row.name);
  }
  return null;
}
