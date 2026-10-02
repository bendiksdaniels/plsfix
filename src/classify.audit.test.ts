// Audit cases for the cell classifier: a structured reference is this
// workbook's own wherever its bracket stands, and only a bracket followed by a
// sheet name and "!" is another workbook.
import { describe, expect, it } from "vitest";

import { classifyCell } from "./classify";

describe("structured references outside a table name", () => {
  it("reads a same-row or bare column reference as this workbook's own", () => {
    expect(classifyCell("=[@Amount]", 4)).toBe("formula");
    expect(classifyCell("=SUM([Amount])", 4)).toBe("formula");
    expect(classifyCell("=[@[Unit price]]*[@Qty]", 4)).toBe("formula");
    expect(classifyCell("=[@Amount]*2", 8)).toBe("partial");
  });

  it("still reads a workbook bracket as external", () => {
    expect(classifyCell("='[Budget.xlsx]Model plan'!$B$4", 4)).toBe("external");
    expect(classifyCell("=[1]Sheet1!A1+[@Amount]", 4)).toBe("external");
    expect(classifyCell("=+[1]S!B1", 4)).toBe("external");
  });

  it("reads a structured reference beside another sheet as cross-sheet", () => {
    expect(classifyCell("=[@Amount]*Sheet2!B1", 4)).toBe("crossSheet");
  });
});
