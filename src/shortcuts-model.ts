// The pure model behind the shortcut manager: public/shortcuts.json's own
// shape read into printable rows, the key grammar Microsoft documents for
// custom shortcuts, and the map Office.actions.replaceShortcuts takes.
// Invariant: what goes to Office is always the Ctrl/Alt/Shift form - the Mac
// names are a display swap only, never what is sent.

/** One remappable action: what shortcuts.json calls it and its shipped key. */
export interface ShortcutRow {
  id: string;
  name: string;
  defaultKey: string;
}

/** The three modifier names Office's own key strings use. */
export type Modifier = "Ctrl" | "Shift" | "Alt";

/** A validated key combination: at least one modifier plus exactly one key. */
export interface Combo {
  modifiers: Modifier[];
  key: string;
}

/** Either the combination, or the one sentence the pane shows instead. */
export type ComboResult =
  { ok: true; combo: Combo } | { ok: false; reason: string };

/** Where the pane is running, for the modifier names it prints. */
export type KeyPlatform = "mac" | "windows";

/**
 * Every refusal says the same thing: the grammar is short enough that naming
 * which of its four rules was broken would be longer than the rule itself.
 */
export const INVALID_COMBO_REASON =
  "Use Ctrl, Alt or Shift plus one key, for example Ctrl+Shift+A.";

// Office writes Ctrl+Shift+Alt+x in its own good-practice pattern, and
// public/shortcuts.json follows it, so that is the order this module prints.
const MODIFIER_ORDER: Modifier[] = ["Ctrl", "Shift", "Alt"];

// The names a typist may reach for. Cmd is the Mac name of Ctrl and Option the
// Mac name of Alt; both are accepted on input and folded to the API's form.
const MODIFIER_NAMES = new Map<string, Modifier>([
  ["ctrl", "Ctrl"],
  ["control", "Ctrl"],
  ["cmd", "Ctrl"],
  ["command", "Ctrl"],
  ["shift", "Shift"],
  ["alt", "Alt"],
  ["option", "Alt"],
  ["opt", "Alt"],
]);

// What the Mac shows instead, per the documented mapping (Alt is Option there)
// plus the Cmd key Office maps Ctrl onto on that platform.
const MAC_NAMES: Record<Modifier, string> = {
  Ctrl: "Cmd",
  Shift: "Shift",
  Alt: "Option",
};

// "Key combinations can include characters A-Z, a-z, 0-9, and the punctuation
// marks -, _ and +" - and "-" and "_" sit on one physical key, so they are the
// same shortcut and fold to one.
const KEY_CHARACTER = /^[A-Za-z0-9\-_+]$/;

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
    if (key !== null || !KEY_CHARACTER.test(part))
      return { ok: false, reason: INVALID_COMBO_REASON };
    key = part === "_" ? "-" : part.toUpperCase();
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

// "+" is both the separator and a legal key, so a trailing one is the key
// rather than an empty part; every other empty part is a typo.
function splitCombo(text: string): string[] | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  const plusKey = trimmed.length > 2 && trimmed.endsWith("++");
  const body = plusKey ? trimmed.slice(0, -2) : trimmed;
  const parts = body.split("+").map((part) => part.trim());
  if (parts.some((part) => part === "")) return null;
  return plusKey ? [...parts, "+"] : parts;
}

/** The API form by default; the Mac names only where the pane asks for them. */
export function formatCombo(
  combo: Combo,
  platform: KeyPlatform = "windows",
): string {
  const names =
    platform === "mac"
      ? combo.modifiers.map((modifier) => MAC_NAMES[modifier])
      : combo.modifiers;
  return [...names, combo.key].join("+");
}

/**
 * A stored key string as the pane prints it. One Office writes in a form this
 * grammar does not cover (an arrow key, say) is shown exactly as it stands
 * rather than swallowed.
 */
export function showKey(key: string, platform: KeyPlatform): string {
  const parsed = parseCombo(key);
  return parsed.ok ? formatCombo(parsed.combo, platform) : key;
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
    const shipped = parseCombo(row.defaultKey);
    const isDefault = shipped.ok && formatCombo(shipped.combo) === wanted;
    map[row.id] = isDefault ? null : wanted;
  }
  return map;
}

/** Combinations the map puts on more than one action, first appearance first. */
export function duplicateKeys(map: Record<string, string | null>): string[] {
  const seen = new Set<string>();
  const repeated: string[] = [];
  for (const key of Object.values(map)) {
    if (key === null) continue;
    if (seen.has(key)) {
      if (!repeated.includes(key)) repeated.push(key);
      continue;
    }
    seen.add(key);
  }
  return repeated;
}
