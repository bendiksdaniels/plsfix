// The pure model behind the shortcut manager: reading the shipped
// public/shortcuts.json into rows, the key grammar Microsoft documents, and the
// map replaceShortcuts takes. No DOM, no Office.js - the panel that drives all
// of this is covered in src/pane/shortcuts-panel*.test.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  diffShortcuts,
  firstClash,
  formatCombo,
  INVALID_COMBO_REASON,
  parseCombo,
  parseShortcutRows,
  showKey,
  type Combo,
  type ShortcutRow,
} from "./shortcuts-model";

function shippedSource(): unknown {
  return JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  );
}

function combo(text: string): Combo {
  const result = parseCombo(text);
  if (!result.ok) throw new Error(`${text} did not parse: ${result.reason}`);
  return result.combo;
}

function reason(text: string): string {
  const result = parseCombo(text);
  if (result.ok) throw new Error(`${text} parsed, expected a refusal`);
  return result.reason;
}

const ROWS: ShortcutRow[] = [
  {
    id: "PLSFIX_SHOWPANE",
    name: "Open pls,fix",
    defaultKey: "Ctrl+Shift+M",
  },
  {
    id: "PLSFIX_AUTOCOLOR",
    name: "Autocolor selection",
    defaultKey: "Ctrl+Shift+K",
  },
  {
    id: "PLSFIX_FILLRIGHT",
    name: "Fill formula right",
    defaultKey: "Ctrl+Alt+R",
  },
];

describe("parseShortcutRows: the shipped shortcuts.json", () => {
  it("reads every binding as an id, a name and its default key", () => {
    const rows = parseShortcutRows(shippedSource());
    expect(rows.length).toBe(42);
    expect(rows[0]).toEqual({
      id: "PLSFIX_SHOWPANE",
      name: "Open pls,fix",
      defaultKey: "Ctrl+Shift+M",
    });
    // The card's order is the shortcuts array's order, not the actions array's.
    expect(rows.map((row) => row.id)).toContain("PLSFIX_STYLES_SCAN");
  });

  it("skips an action nothing binds a key to", () => {
    const rows = parseShortcutRows({
      actions: [
        { id: "A", name: "Alpha" },
        { id: "B", name: "Beta" },
      ],
      shortcuts: [{ action: "A", key: { default: "Ctrl+Shift+A" } }],
    });
    expect(rows).toEqual([
      { id: "A", name: "Alpha", defaultKey: "Ctrl+Shift+A" },
    ]);
  });

  it("names the stage when the file is not a shortcuts file", () => {
    expect(() => parseShortcutRows({ nope: true })).toThrow(
      /parse shortcuts\.json/,
    );
    expect(() => parseShortcutRows(null)).toThrow(/parse shortcuts\.json/);
  });

  it("names the binding whose action is missing", () => {
    expect(() =>
      parseShortcutRows({
        actions: [],
        shortcuts: [{ action: "GHOST", key: { default: "Ctrl+Shift+A" } }],
      }),
    ).toThrow(/GHOST/);
  });
});

describe("parseCombo: the grammar Office documents", () => {
  it("takes one modifier and one key", () => {
    expect(combo("Ctrl+Shift+A")).toEqual({
      modifiers: ["Ctrl", "Shift"],
      key: "A",
    });
    expect(combo("Alt+1")).toEqual({ modifiers: ["Alt"], key: "1" });
  });

  it("writes the modifiers in Office's own order whatever order they came in", () => {
    expect(formatCombo(combo("Alt+shift+ctrl+8"))).toBe("Ctrl+Shift+Alt+8");
  });

  // The docs say Cmd is supported on macOS and map Cmd to Ctrl on Windows, so
  // Cmd is a modifier of its own and is passed to Office as written.
  it("keeps Cmd as its own modifier and folds only its long spelling", () => {
    expect(formatCombo(combo("Cmd+Shift+A"))).toBe("Cmd+Shift+A");
    expect(formatCombo(combo("Command+Shift+A"))).toBe("Cmd+Shift+A");
    expect(combo("Cmd+Ctrl+A").modifiers).toEqual(["Ctrl", "Cmd"]);
  });

  // "On macOS, the Alt key is mapped to the Option key": one modifier, two names.
  it("folds Option onto Alt, the documented equivalence", () => {
    expect(formatCombo(combo("Option+P"))).toBe("Alt+P");
    expect(formatCombo(combo("Cmd+Opt+P"))).toBe("Cmd+Alt+P");
  });

  it("uppercases the letter and folds the two synonyms of one physical key", () => {
    expect(combo("ctrl+a").key).toBe("A");
    expect(formatCombo(combo("Ctrl+_"))).toBe("Ctrl+-");
  });

  it("takes the three punctuation keys the docs list, + included", () => {
    expect(combo("Ctrl+Shift++").key).toBe("+");
    expect(combo("Ctrl + +").key).toBe("+");
    expect(combo("Ctrl+-").key).toBe("-");
  });

  // Microsoft's own manifest samples bind Ctrl+Alt+Up and Ctrl+Alt+Down.
  it("takes the four arrow keys the docs' own samples use", () => {
    expect(formatCombo(combo("Ctrl+Alt+Up"))).toBe("Ctrl+Alt+Up");
    expect(formatCombo(combo("ctrl+alt+down"))).toBe("Ctrl+Alt+Down");
    expect(formatCombo(combo("Ctrl+Alt+left"))).toBe("Ctrl+Alt+Left");
    expect(formatCombo(combo("Ctrl+Alt+RIGHT"))).toBe("Ctrl+Alt+Right");
  });

  it("refuses a combination with no modifier, or with Shift as the only one", () => {
    expect(reason("A")).toBe(INVALID_COMBO_REASON);
    expect(reason("Shift+A")).toBe(INVALID_COMBO_REASON);
  });

  it("refuses two keys, no key, a repeated modifier and an unknown key", () => {
    expect(reason("Ctrl+A+B")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Shift")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Ctrl+A")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Home")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+++")).toBe(INVALID_COMBO_REASON);
    expect(reason("")).toBe(INVALID_COMBO_REASON);
    expect(reason("   ")).toBe(INVALID_COMBO_REASON);
  });

  it("ignores the spaces a typist leaves around the plus", () => {
    expect(formatCombo(combo(" Ctrl + Shift + A "))).toBe("Ctrl+Shift+A");
  });
});

describe("showKey: what the pane prints", () => {
  it("prints a key exactly as Office takes it, on every platform", () => {
    expect(showKey("ctrl+shift+a")).toBe("Ctrl+Shift+A");
    expect(showKey("Cmd+Shift+A")).toBe("Cmd+Shift+A");
  });

  it("hands back a key it cannot parse unchanged rather than hiding it", () => {
    expect(showKey("Ctrl+Alt+PgUp")).toBe("Ctrl+Alt+PgUp");
  });
});

describe("diffShortcuts: the map replaceShortcuts takes", () => {
  it("sends null for a box left empty", () => {
    expect(diffShortcuts(ROWS, {})).toEqual({
      PLSFIX_SHOWPANE: null,
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
    expect(diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "  " })).toEqual({
      PLSFIX_SHOWPANE: null,
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("sends the canonical combination for a box that differs", () => {
    expect(diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "alt+ctrl+j" })).toEqual({
      PLSFIX_SHOWPANE: null,
      PLSFIX_AUTOCOLOR: "Ctrl+Alt+J",
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("sends null for a box retyping the default, however it was written", () => {
    expect(
      diffShortcuts(ROWS, {
        PLSFIX_AUTOCOLOR: "ctrl+shift+k",
        PLSFIX_FILLRIGHT: " Alt + Ctrl + r ",
      }),
    ).toEqual({
      PLSFIX_SHOWPANE: null,
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("ignores a box for an action the file no longer lists", () => {
    expect(diffShortcuts(ROWS, { PLSFIX_GONE: "Ctrl+Shift+Z" })).toEqual({
      PLSFIX_SHOWPANE: null,
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("names the row when a value the panel should have validated is not one", () => {
    expect(() => diffShortcuts(ROWS, { PLSFIX_FILLRIGHT: "Shift+R" })).toThrow(
      /PLSFIX_FILLRIGHT/,
    );
  });
});

describe("firstClash: two actions on one key", () => {
  it("finds nothing when every effective key is its own", () => {
    expect(firstClash(ROWS, diffShortcuts(ROWS, {}))).toBeNull();
    expect(
      firstClash(ROWS, diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "Ctrl+Alt+J" })),
    ).toBeNull();
  });

  it("catches a custom key that lands on another action's SHIPPED default", () => {
    const map = diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "Ctrl+Shift+M" });
    expect(firstClash(ROWS, map)).toEqual({
      key: "Ctrl+Shift+M",
      first: "Open pls,fix",
      second: "Autocolor selection",
    });
  });

  it("catches one combination typed into two boxes", () => {
    const map = diffShortcuts(ROWS, {
      PLSFIX_AUTOCOLOR: "Ctrl+Alt+J",
      PLSFIX_FILLRIGHT: "ctrl+alt+j",
    });
    expect(firstClash(ROWS, map)).toEqual({
      key: "Ctrl+Alt+J",
      first: "Autocolor selection",
      second: "Fill formula right",
    });
  });

  it("reports the pair the reader meets first, in row order", () => {
    const map = diffShortcuts(ROWS, {
      PLSFIX_AUTOCOLOR: "Ctrl+Alt+R",
      PLSFIX_FILLRIGHT: "Ctrl+Shift+M",
    });
    expect(firstClash(ROWS, map)?.first).toBe("Open pls,fix");
    expect(firstClash(ROWS, map)?.key).toBe("Ctrl+Shift+M");
  });

  it("frees a default the action that owned it has moved off", () => {
    const map = diffShortcuts(ROWS, {
      PLSFIX_SHOWPANE: "Ctrl+Alt+P",
      PLSFIX_AUTOCOLOR: "Ctrl+Shift+M",
    });
    expect(firstClash(ROWS, map)).toBeNull();
  });
});
