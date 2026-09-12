// The pure model behind the shortcut manager: reading the shipped
// public/shortcuts.json into rows, the key grammar Office documents, the Mac
// display swap and the map replaceShortcuts takes. No DOM, no Office.js - the
// panel that drives all of this is covered in src/pane/shortcuts-panel.test.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  diffShortcuts,
  duplicateKeys,
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

  it("takes Cmd and Option as the Mac names of Ctrl and Alt", () => {
    expect(formatCombo(combo("Cmd+Shift+A"))).toBe("Ctrl+Shift+A");
    expect(formatCombo(combo("Command+Option+P"))).toBe("Ctrl+Alt+P");
  });

  it("uppercases the letter and folds the two synonyms of one physical key", () => {
    expect(combo("ctrl+a").key).toBe("A");
    expect(formatCombo(combo("Ctrl+_"))).toBe("Ctrl+-");
  });

  it("takes the three punctuation keys the docs list, + included", () => {
    expect(combo("Ctrl+Shift++").key).toBe("+");
    expect(combo("Ctrl+-").key).toBe("-");
  });

  it("refuses a combination with no modifier, or with Shift as the only one", () => {
    expect(reason("A")).toBe(INVALID_COMBO_REASON);
    expect(reason("Shift+A")).toBe(INVALID_COMBO_REASON);
  });

  it("refuses two keys, no key, a repeated modifier and an unknown key", () => {
    expect(reason("Ctrl+A+B")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Shift")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Ctrl+A")).toBe(INVALID_COMBO_REASON);
    expect(reason("Ctrl+Up")).toBe(INVALID_COMBO_REASON);
    expect(reason("")).toBe(INVALID_COMBO_REASON);
    expect(reason("   ")).toBe(INVALID_COMBO_REASON);
  });

  it("ignores the spaces a typist leaves around the plus", () => {
    expect(formatCombo(combo(" Ctrl + Shift + A "))).toBe("Ctrl+Shift+A");
  });
});

describe("formatCombo and showKey: what the pane shows", () => {
  it("writes the API form off a Mac", () => {
    expect(formatCombo(combo("Ctrl+Shift+A"), "windows")).toBe("Ctrl+Shift+A");
    expect(showKey("Ctrl+Shift+A", "windows")).toBe("Ctrl+Shift+A");
  });

  it("shows Cmd and Option on a Mac while the API still takes Ctrl and Alt", () => {
    expect(formatCombo(combo("Ctrl+Shift+Alt+8"), "mac")).toBe(
      "Cmd+Shift+Option+8",
    );
    expect(showKey("Ctrl+Alt+R", "mac")).toBe("Cmd+Option+R");
    // The value that goes to Office is untouched by the display swap.
    expect(formatCombo(combo("Ctrl+Alt+R"))).toBe("Ctrl+Alt+R");
  });

  it("hands back a key it cannot parse unchanged rather than hiding it", () => {
    expect(showKey("Ctrl+Alt+Up", "mac")).toBe("Ctrl+Alt+Up");
  });
});

describe("diffShortcuts: the map replaceShortcuts takes", () => {
  it("sends null for a box left empty", () => {
    expect(diffShortcuts(ROWS, {})).toEqual({
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
    expect(diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "  " })).toEqual({
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("sends the canonical combination for a box that differs", () => {
    expect(diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "alt+ctrl+j" })).toEqual({
      PLSFIX_AUTOCOLOR: "Ctrl+Alt+J",
      PLSFIX_FILLRIGHT: null,
    });
  });

  it("sends null for a box retyping the default, however it was written", () => {
    expect(
      diffShortcuts(ROWS, {
        PLSFIX_AUTOCOLOR: "ctrl+shift+k",
        PLSFIX_FILLRIGHT: "Cmd+Option+R",
      }),
    ).toEqual({ PLSFIX_AUTOCOLOR: null, PLSFIX_FILLRIGHT: null });
  });

  it("ignores a box for an action the file no longer lists", () => {
    expect(diffShortcuts(ROWS, { PLSFIX_GONE: "Ctrl+Shift+Z" })).toEqual({
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

describe("duplicateKeys: one combination on two actions", () => {
  it("finds nothing in a map with no repeat", () => {
    expect(
      duplicateKeys({ a: "Ctrl+Shift+A", b: "Ctrl+Shift+B", c: null }),
    ).toEqual([]);
  });

  it("names each repeated combination once, in the order it first appears", () => {
    expect(
      duplicateKeys({
        a: "Ctrl+Shift+A",
        b: "Ctrl+Shift+B",
        c: "Ctrl+Shift+A",
        d: "Ctrl+Shift+B",
        e: "Ctrl+Shift+A",
      }),
    ).toEqual(["Ctrl+Shift+A", "Ctrl+Shift+B"]);
  });

  it("does not count the nulls as a repeat", () => {
    expect(duplicateKeys({ a: null, b: null })).toEqual([]);
  });
});
