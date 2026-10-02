// Pass-1 hunt: src/cycle-labels.ts stood at 50 % branches with no direct unit
// test at all - only the toast text of a live cycle press exercised it, and
// every real press only ever hands alignLabel/underlineLabel a value already
// in their own map (nextAlignment/nextUnderline in src/cycles.ts never answer
// anything else). The fallback arm is a genuine second branch of a pure,
// exported function, so it gets its own pin rather than staying dark.
import { describe, expect, it } from "vitest";
import {
  alignLabel,
  numberCycleLabels,
  underlineLabel,
} from "../../src/cycle-labels";
import { DEFAULT_SETTINGS } from "../../src/settings";

describe("alignLabel", () => {
  it("names the four ladder rungs in lower case", () => {
    expect(alignLabel("Left")).toBe("left");
    expect(alignLabel("Center")).toBe("centre");
    expect(alignLabel("Right")).toBe("right");
    expect(alignLabel("General")).toBe("general");
  });

  it("falls back to a lower-cased echo for a value it does not name", () => {
    // Unreachable through the format cycle today - nextAlignment only ever
    // steps to one of the four rungs above - but the function is exported and
    // pure, and a caller outside the cycle must get a sentence back, not
    // undefined.
    expect(alignLabel("Justify")).toBe("justify");
    expect(alignLabel("Distributed")).toBe("distributed");
  });
});

describe("underlineLabel", () => {
  it("names the three ladder rungs in lower case", () => {
    expect(underlineLabel("Single")).toBe("single");
    expect(underlineLabel("Double")).toBe("double");
    expect(underlineLabel("None")).toBe("none");
  });

  it("falls back to a lower-cased echo for a value it does not name", () => {
    // Unreachable through the format cycle today for the same reason: the
    // fallback is what keeps an unnamed underline from printing undefined
    // rather than a sentence, should a future caller ever feed it one.
    expect(underlineLabel("SingleAccountant")).toBe("singleaccountant");
    expect(underlineLabel("Slant")).toBe("slant");
  });
});

describe("numberCycleLabels", () => {
  it("places the currency symbol per language, before or after the digits", () => {
    const lv = numberCycleLabels({ ...DEFAULT_SETTINGS, language: "lv" });
    const en = numberCycleLabels({ ...DEFAULT_SETTINGS, language: "en" });
    expect(lv.currency[0]).toBe(`1,234 ${DEFAULT_SETTINGS.currency}`);
    expect(en.currency[0]).toBe(`${DEFAULT_SETTINGS.currency} 1,234`);
    // Every other family is static, and the same regardless of language.
    expect(lv.date).toEqual(en.date);
    expect(lv.percent).toEqual(en.percent);
  });

  it("names the third currency rung as thousands, matching the cycle's own scaling", () => {
    const labels = numberCycleLabels(DEFAULT_SETTINGS);
    expect(labels.currency[2]).toContain("(thousands)");
  });
});
