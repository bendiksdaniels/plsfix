// The house number style against what Excel actually hands back: the
// separators the host reports come out of the operating system's locale data,
// not out of a keyboard, so the comparison has to canonicalise both sides.

import { describe, expect, it } from "vitest";

import { formatAmount, separatorSample, separatorsMatch } from "./numbers";

// Excel.Application.thousandsSeparator is "based on Excel's local settings"
// (learn.microsoft.com/en-us/javascript/api/excel/excel.application), and the
// Latvian and Russian locales spell that setting as U+00A0 NO-BREAK SPACE in
// both Microsoft's NLS data and CLDR. Every one of these shows on screen as
// the gap the house style asks for.
const NBSP = "\u00A0";
const NARROW_NBSP = "\u202F";
const THIN_SPACE = "\u2009";

describe("separators as the host reports them", () => {
  it("counts a non-breaking thousands space as the house space", () => {
    expect(separatorsMatch(".", NBSP, "lv")).toBe(true);
    expect(separatorsMatch(".", NBSP, "ru")).toBe(true);
  });

  it("counts the narrow and thin spaces the same way", () => {
    expect(separatorsMatch(".", NARROW_NBSP, "lv")).toBe(true);
    expect(separatorsMatch(".", THIN_SPACE, "ru")).toBe(true);
  });

  it("still refuses a separator that is not a space at all", () => {
    expect(separatorsMatch(".", ",", "lv")).toBe(false);
    expect(separatorsMatch(",", " ", "lv")).toBe(false);
    expect(separatorsMatch(".", " ", "en")).toBe(false);
    expect(separatorsMatch(".", NBSP, "en")).toBe(false);
  });

  // The sample stays verbatim: it is what Excel shows, and a non-breaking
  // space renders as the same gap.
  it("samples with the separator the host reported", () => {
    expect(separatorSample(".", NBSP)).toBe(`1${NBSP}094${NBSP}417.5`);
  });
});

describe("amounts at the edges", () => {
  it("keeps zero, tiny negatives and round thousands readable", () => {
    expect(formatAmount(0, "lv")).toBe("0");
    expect(formatAmount(-0.04, "lv")).toBe("0");
    expect(formatAmount(999.99, "lv")).toBe("1 000");
    expect(formatAmount(-999.99, "en")).toBe("-1,000");
  });

  it("groups a number with no decimals asked for", () => {
    expect(formatAmount(1234567, "lv", 0)).toBe("1 234 567");
    expect(formatAmount(1234567, "en", 0)).toBe("1,234,567");
  });
});
