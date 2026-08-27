import { describe, expect, it } from "vitest";
import { classifyCell } from "./classify";

describe("constants and empty cells", () => {
  it("calls empty cells blank", () => {
    expect(classifyCell(null, null)).toBe("blank");
    expect(classifyCell("", "")).toBe("blank");
  });

  it("calls typed content an input", () => {
    expect(classifyCell(25, 25)).toBe("input");
    expect(classifyCell("Revenue", "Revenue")).toBe("input");
    expect(classifyCell(true, true)).toBe("input");
    expect(classifyCell(0, 0)).toBe("input");
  });

  it("keeps a formula that returns nothing a formula", () => {
    expect(classifyCell("=A1", "")).toBe("formula");
  });
});

describe("same-sheet formulas", () => {
  it("calls reference-only formulas formulas", () => {
    expect(classifyCell("=A1+B1", 3)).toBe("formula");
    expect(classifyCell("=SUM(A1:A9)", 9)).toBe("formula");
    expect(classifyCell("=DAYS360($A$1,$B$12)", 30)).toBe("formula");
  });

  it("does not read digits in function names as hardcodes", () => {
    expect(classifyCell("=LOG10(A1)", 1)).toBe("formula");
    expect(classifyCell("=TODAY()", 46000)).toBe("formula");
  });
});

describe("links", () => {
  it("flags references to another sheet", () => {
    expect(classifyCell("=SUM(Sheet2!A1:A9)", 9)).toBe("crossSheet");
    expect(classifyCell("='Assumptions 2026'!B4", 1)).toBe("crossSheet");
  });

  it("flags references to another workbook", () => {
    expect(classifyCell("=[Model.xlsx]Sheet1!A1", 1)).toBe("external");
    expect(classifyCell("='[Budget 2026.xlsx]Sheet1'!A1", 1)).toBe("external");
  });

  it("ranks external above cross-sheet above hardcodes", () => {
    expect(classifyCell("=[Book1.xlsx]Sheet1!A1*1.05", 1)).toBe("external");
    expect(classifyCell("=Sheet2!A1*1.05", 1)).toBe("crossSheet");
  });
});

describe("partial inputs", () => {
  it("flags numbers typed inside a formula", () => {
    expect(classifyCell("=A1*1.05", 1)).toBe("partial");
    expect(classifyCell("=A1+100", 101)).toBe("partial");
    expect(classifyCell("=A1*.5", 0.5)).toBe("partial");
    expect(classifyCell("=A1*5%", 0.05)).toBe("partial");
  });

  // Macabacus counts a comparison against a constant as a hardcode too: the
  // threshold is an assumption that belongs in a cell.
  it("counts constants in comparisons", () => {
    expect(classifyCell('=IF(A1>0,"over 100",B1)', "over 100")).toBe("partial");
  });
});

describe("string literals", () => {
  it("ignores digits, brackets and bangs inside quoted text", () => {
    expect(classifyCell('=A1&" 2026"', "x 2026")).toBe("formula");
    expect(classifyCell('=A1&"[note]"', "x[note]")).toBe("formula");
    expect(classifyCell('=A1&"wow!"', "wow!")).toBe("formula");
  });

  it("handles doubled quotes inside a literal", () => {
    expect(classifyCell('="say ""99"" now"&A1', 'say "99" now')).toBe("formula");
  });

  it("survives an unterminated literal", () => {
    expect(classifyCell('="abc', "abc")).toBe("formula");
  });
});
