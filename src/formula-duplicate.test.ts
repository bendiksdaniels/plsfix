import { describe, expect, it } from "vitest";
import { duplicateFormula, duplicateFormulas } from "./formula-duplicate";

// The copied block: Model!A1:C3, pasted four rows down at Model!A5:C7.
const BLOCK = {
  sheet: "Model",
  row: 0,
  column: 0,
  rowCount: 3,
  columnCount: 3,
};
const DOWN = { rows: 4, columns: 0 };

describe("duplicateFormula onto the same sheet", () => {
  it.each([
    // Inside the block: the reference moves with it, markers as written.
    ["=A1+1", "=A5+1"],
    ["=$A$1", "=$A$5"],
    ["=A$1", "=A$5"],
    ["=$A1", "=$A5"],
    ["=SUM(A1:C3)", "=SUM(A5:C7)"],
    ["=SUM(A1:A3)", "=SUM(A5:A7)"],
    ["=SUM($A$1:$C$3)", "=SUM($A$5:$C$7)"],
    ["=IF(A1>B2,C3,Z9)", "=IF(A5>B6,C7,Z9)"],
    ['=A1&"ok"&B2', '=A5&"ok"&B6'],
    // Outside the block: same cells, markers as written.
    ["=Z9", "=Z9"],
    ["=$Z$9", "=$Z$9"],
    ["=A4", "=A4"],
    ["=D1", "=D1"],
    // A range straddling the edge points partly outside, so it stays put.
    ["=SUM(A1:C4)", "=SUM(A1:C4)"],
    // Whole columns and whole rows span the sheet, never a copied block.
    ["=SUM(A:A)", "=SUM(A:A)"],
    ["=SUM($A:$A)", "=SUM($A:$A)"],
    ["=SUM(1:1)", "=SUM(1:1)"],
    // The block's own sheet by name is inside; any other sheet is outside.
    ["=Model!B2", "=Model!B6"],
    ["=model!B2", "=model!B6"],
    ["='Model'!A1", "='Model'!A5"],
    ["=Data!B2", "=Data!B2"],
    ["='P&L 2025'!A1:B2", "='P&L 2025'!A1:B2"],
    ["=[Book1.xlsx]Model!A1", "=[Book1.xlsx]Model!A1"],
    // Text, tables, names and function names are not references.
    ['="A1"&A1', '="A1"&A5'],
    ["=SUM(Table1[Col])", "=SUM(Table1[Col])"],
    ["=Tax_A1", "=Tax_A1"],
    ["=LOG10(A1)", "=LOG10(A5)"],
    ["=A1+Sheet2!A1+$C$3", "=A5+Sheet2!A1+$C$7"],
  ])("rewrites %s as %s", (formula, expected) => {
    expect(duplicateFormula(formula, BLOCK, DOWN)).toBe(expected);
  });

  it("treats a destination sheet of the same name as the same sheet", () => {
    expect(duplicateFormula("=Z9", BLOCK, DOWN, "model")).toBe("=Z9");
    expect(duplicateFormula("=A1", BLOCK, DOWN, "Model")).toBe("=A5");
  });

  it("shifts columns as well as rows", () => {
    expect(duplicateFormula("=B2", BLOCK, { rows: 0, columns: 3 })).toBe("=E2");
    expect(duplicateFormula("=$B$2", BLOCK, { rows: 1, columns: 3 })).toBe(
      "=$E$3",
    );
  });

  it("leaves a reference that would fall off the grid where it is", () => {
    expect(duplicateFormula("=A1", BLOCK, { rows: -1, columns: 0 })).toBe(
      "=A1",
    );
    expect(duplicateFormula("=A1", BLOCK, { rows: 0, columns: -1 })).toBe(
      "=A1",
    );
  });

  it("never touches a cell that is not a formula", () => {
    expect(duplicateFormula("A1", BLOCK, DOWN)).toBe("A1");
    expect(duplicateFormula("", BLOCK, DOWN)).toBe("");
  });
});

// Onto another sheet an outside reference only keeps pointing at the SAME
// cells once it names the sheet it was written on, and an in-block reference
// written with the source sheet's name has to follow the block instead.
describe("duplicateFormula onto another sheet", () => {
  it.each([
    // Inside, unqualified: still the block's own cell, still unqualified.
    ["=A1+10", "=A5+10"],
    ["=SUM(A1:C3)", "=SUM(A5:C7)"],
    ["=$A$1", "=$A$5"],
    // Outside, unqualified: qualified with the source sheet, address as written.
    ["=Z9", "=Model!Z9"],
    ["=$Z$9", "=Model!$Z$9"],
    ["=Z9*B1", "=Model!Z9*B5"],
    ["=SUM(A:A)", "=SUM(Model!A:A)"],
    ["=SUM(1:1)", "=SUM(Model!1:1)"],
    ["=SUM(A1:C4)", "=SUM(Model!A1:C4)"],
    // Outside, already qualified with the source sheet: byte for byte.
    ["=Model!Z9", "=Model!Z9"],
    ["='Model'!Z9", "='Model'!Z9"],
    // Inside, qualified with the source sheet: the destination's own cell.
    ["=Model!B2", "=Data!B6"],
    ["='Model'!B2", "=Data!B6"],
    // Another sheet stays another sheet, the destination included.
    ["=Data!B2", "=Data!B2"],
    ["=Sheet2!A1", "=Sheet2!A1"],
    ["=[Book1.xlsx]Model!A1", "=[Book1.xlsx]Model!A1"],
    // Text, tables and names are still not references.
    ['="A1"&Z9', '="A1"&Model!Z9'],
    ["=SUM(Table1[Col])", "=SUM(Table1[Col])"],
    ["=Tax_A1", "=Tax_A1"],
    ["=LOG10(A1)", "=LOG10(A5)"],
  ])("rewrites %s as %s", (formula, expected) => {
    expect(duplicateFormula(formula, BLOCK, DOWN, "Data")).toBe(expected);
  });

  it("quotes a source sheet name that needs quoting", () => {
    const block = { ...BLOCK, sheet: "P&L 2025" };
    expect(duplicateFormula("=Z9", block, DOWN, "Data")).toBe("='P&L 2025'!Z9");
    expect(
      duplicateFormula("=Z9", { ...BLOCK, sheet: "Q1" }, DOWN, "Data"),
    ).toBe("='Q1'!Z9");
  });

  it("quotes a destination sheet name that needs quoting", () => {
    expect(duplicateFormula("=Model!B2", BLOCK, DOWN, "P&L 2025")).toBe(
      "='P&L 2025'!B6",
    );
  });

  it("leaves everything alone when either sheet name is unknown", () => {
    expect(duplicateFormula("=Z9", { ...BLOCK, sheet: "" }, DOWN, "Data")).toBe(
      "=Z9",
    );
    expect(duplicateFormula("=Z9", BLOCK, DOWN, "")).toBe("=Z9");
  });
});

describe("duplicateFormulas", () => {
  it("rewrites a grid and passes plain values through", () => {
    expect(
      duplicateFormulas(
        [
          [1, "=A1*2"],
          ["label", "=SUM($A$1:$C$3)+Z9"],
        ],
        BLOCK,
        DOWN,
      ),
    ).toEqual([
      [1, "=A5*2"],
      ["label", "=SUM($A$5:$C$7)+Z9"],
    ]);
  });

  it("carries the destination sheet into every cell", () => {
    expect(duplicateFormulas([[1, "=A1*2+Z9"]], BLOCK, DOWN, "Data")).toEqual([
      [1, "=A5*2+Model!Z9"],
    ]);
  });
});
