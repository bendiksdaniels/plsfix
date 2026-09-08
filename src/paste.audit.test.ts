// Slice B audit: the pure edges of the fill/format/formula maths the Excel
// adapter drives - a formula that never closes its bracket, a number format
// whose separator is backslash-escaped, an IFERROR with no fallback, and grids
// the host hands back ragged.
import { describe, expect, it } from "vitest";
import { classifyCell } from "./classify";
import { analyzeGrid } from "./model";
import {
  detectFillExtent,
  flipSign,
  formatDecimals,
  stepDecimals,
  toggleIfError,
} from "./paste";

describe("a formula that never closes", () => {
  it("wraps rather than unwraps what it cannot match", () => {
    // Only our own "=-(...)" wrap unwraps; a bracket with no partner is not it.
    expect(flipSign([["=-(A1"]])).toEqual([["=-(-(A1)"]]);
  });

  it("takes an unclosed table bracket for what it is", () => {
    expect(classifyCell("=Table1[", 1)).toBe("formula");
  });
});

describe("the IFERROR guard", () => {
  it("guards a formula that opens with + or @", () => {
    expect(toggleIfError([["=+A1/B1", "=@A1:A5"]], "0")).toEqual([
      ["=IFERROR(+A1/B1,0)", "=IFERROR(@A1:A5,0)"],
    ]);
  });

  it("strips a guard whose own fallback holds brackets", () => {
    expect(toggleIfError([['=IFERROR(IF(A1,1),"n/a")']], "0")).toEqual([
      ["=IF(A1,1)"],
    ]);
  });

  it("leaves an IFERROR with no fallback alone and guards it", () => {
    // No comma at depth 1: this is not the guard this button writes.
    expect(toggleIfError([["=IFERROR(A1)"]], "0")).toEqual([
      ["=IFERROR(IFERROR(A1),0)"],
    ]);
  });

  it("passes an array constant through untouched", () => {
    expect(toggleIfError([["={1,2,3}"]], "0")).toEqual([
      ["=IFERROR({1,2,3},0)"],
    ]);
  });
});

describe("a number format with backslash escapes", () => {
  it("steps the digits, not the escaped separator", () => {
    expect(stepDecimals('#,##0\\ "kr"', 1)).toBe('#,##0.0\\ "kr"');
  });

  it("reads an escaped percent sign as a character, not a percentage", () => {
    // "0.0\%" prints a % without the hundredfold shift, so it rounds to one
    // decimal where "0.0%" rounds the stored value to three.
    expect(formatDecimals("0.0\\%")).toBe(1);
    expect(formatDecimals("0.0%")).toBe(3);
  });

  it("counts a thousands comma off the precision", () => {
    expect(formatDecimals("#,##0,,")).toBe(-6);
    expect(stepDecimals("#,##0,,", 1)).toBe("#,##0.0,,");
  });
});

describe("the neighbour lines that size a fill", () => {
  it("is nothing when the cell beside the origin is blank", () => {
    expect(detectFillExtent([["", "a", "b"]])).toBe(0);
    expect(detectFillExtent([])).toBe(0);
  });

  it("takes the longest unbroken run of the two lines", () => {
    expect(detectFillExtent([["a"], ["a", "b", "c"]])).toBe(3);
  });
});

describe("a grid the host hands back ragged", () => {
  it("counts the formulas it has and the values it does not", () => {
    expect(analyzeGrid([["=A1"], ["=A2", "=B2"]], [[1]])).toEqual({
      cells: 3,
      formulas: 3,
      errors: 0,
      blanks: 2,
    });
  });
});
