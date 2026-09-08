import { describe, expect, it } from "vitest";
import { rankTools, type ToolEntry } from "./tool-search";

// A small, hand-built catalogue standing in for the real one: enough to
// prove the four-tier match and the scoring weights without booting a DOM.
// The real catalogue (src/pane/tool-search.ts) is covered by its own test.
const TOOLS: ToolEntry[] = [
  {
    action: "chart-waterfall",
    label: "Waterfall from selection",
    tab: "Tools",
    sentence:
      "Builds a bridge chart from a label and value table. First and last rows are the opening and closing totals.",
  },
  {
    action: "template-ebitda-bridge",
    label: "EBITDA bridge",
    tab: "Tools",
    sentence:
      "Writes an EBITDA bridge table shaped for the Waterfall button: opening, five steps, closing and a check row.",
  },
  {
    action: "insert-color-key",
    label: "Insert color key",
    tab: "Tools",
    sentence:
      "Drops the color legend on the sheet, starting at the active cell.",
  },
  {
    action: "autocolor",
    label: "Autocolor selection",
    tab: "Tools",
    sentence:
      "Colors the selection by content: typed numbers, formulas, cross-sheet links, external links and hardcodes.",
  },
  {
    action: "undo",
    label: "Undo last pls,fix action",
    tab: "Tools",
    sentence:
      "Puts back what the last pls,fix action overwrote. Excel's own Ctrl+Z never sees add-in writes.",
  },
  {
    action: "export-selection",
    label: "Export selection",
    tab: "Links",
    sentence:
      "Sends the selected range as a linked picture. The selection must be one unbroken block.",
  },
];

describe("rankTools", () => {
  it('"wat" finds Waterfall from selection first', () => {
    const results = rankTools("wat", TOOLS);
    expect(results[0]?.label).toBe("Waterfall from selection");
  });

  it('"color key" finds Insert color key above autocolor', () => {
    const results = rankTools("color key", TOOLS);
    const colorKey = results.findIndex((r) => r.label === "Insert color key");
    const autocolor = results.findIndex(
      (r) => r.label === "Autocolor selection",
    );
    expect(colorKey).toBeGreaterThanOrEqual(0);
    // autocolor's label and sentence never say "key", so a two-word AND
    // search leaves it out entirely - which is "above" it in the plainest
    // sense. If a future sentence ever adds the word, this still holds.
    expect(autocolor === -1 || colorKey < autocolor).toBe(true);
  });

  it('"undo" finds the undo button', () => {
    const results = rankTools("undo", TOOLS);
    expect(results[0]?.action).toBe("undo");
  });

  it('"zzz" finds nothing', () => {
    expect(rankTools("zzz", TOOLS)).toEqual([]);
  });

  it("an empty or whitespace query returns no results", () => {
    expect(rankTools("", TOOLS)).toEqual([]);
    expect(rankTools("   ", TOOLS)).toEqual([]);
  });

  it("respects the limit", () => {
    // Every entry above carries "e" somewhere in its label or sentence.
    const results = rankTools("e", TOOLS, 2);
    expect(results).toHaveLength(2);
  });

  it("defaults the limit to 8", () => {
    const many: ToolEntry[] = Array.from({ length: 12 }, (_, i) => ({
      action: `tool-${String(i)}`,
      label: `Sample tool ${String(i)}`,
      tab: "Tools",
      sentence: "A sample entry used only to pad the list past the cap.",
    }));
    expect(rankTools("sample", many)).toHaveLength(8);
  });

  it("excludes a tool missing even one of several query words", () => {
    // "waterfall" alone also hits the bridge template's sentence ("shaped
    // for the Waterfall button") and "selection" alone also hits Autocolor
    // selection and Export selection - but nothing else in the catalogue
    // carries both words, so only the chart itself survives the AND filter.
    const results = rankTools("waterfall selection", TOOLS);
    expect(results.map((r) => r.action)).toEqual(["chart-waterfall"]);
  });

  // hitOf's last fallback (a token buried mid-word in the sentence, not a
  // word-prefix anywhere): "codes" is not a prefix of any label or sentence
  // word, but it is a plain substring of "hardcodes" in autocolor's sentence.
  it('a token mid-word in a sentence ("codes" inside "hardcodes") still counts as a sentence hit', () => {
    const results = rankTools("codes", TOOLS);
    expect(results.map((r) => r.action)).toEqual(["autocolor"]);
  });

  it("keeps ties in the input order", () => {
    const tied: ToolEntry[] = [
      {
        action: "b",
        label: "Beta tool",
        tab: "Tools",
        sentence: "About beta.",
      },
      {
        action: "a",
        label: "Alpha tool",
        tab: "Tools",
        sentence: "About alpha.",
      },
    ];
    // Neither label nor sentence carries "tool" as anything but a shared,
    // identically-weighted word, so both score the same and order is kept.
    const results = rankTools("tool", tied);
    expect(results.map((r) => r.action)).toEqual(["b", "a"]);
  });
});
