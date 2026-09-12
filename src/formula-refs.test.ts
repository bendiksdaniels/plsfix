import { describe, expect, it } from "vitest";
import { findReferences, rewriteReferences } from "./formula-refs";

// The token as written, which is what a rewrite has to be able to put back.
function texts(formula: string): string[] {
  return findReferences(formula).map((ref) => ref.text);
}

describe("findReferences", () => {
  it("reads a bare cell", () => {
    const [ref] = findReferences("=A1");
    expect(ref).toMatchObject({
      text: "A1",
      sheet: "",
      start: 1,
      end: 3,
      pair: false,
    });
    expect(ref?.from).toEqual({
      column: 0,
      columnAbsolute: false,
      row: 0,
      rowAbsolute: false,
    });
  });

  it("keeps every mix of absolute markers", () => {
    expect(findReferences("=$A$1+A$1+$A1").map((ref) => ref.from)).toEqual([
      { column: 0, columnAbsolute: true, row: 0, rowAbsolute: true },
      { column: 0, columnAbsolute: false, row: 0, rowAbsolute: true },
      { column: 0, columnAbsolute: true, row: 0, rowAbsolute: false },
    ]);
  });

  it("reads a range as one reference", () => {
    const [ref] = findReferences("=SUM(A1:B2)");
    expect(ref).toMatchObject({ text: "A1:B2", pair: true });
    expect(ref?.to).toEqual({
      column: 1,
      columnAbsolute: false,
      row: 1,
      rowAbsolute: false,
    });
  });

  it("reads a whole column and a whole row", () => {
    const [column] = findReferences("=SUM($A:$A)");
    expect(column).toMatchObject({ text: "$A:$A", pair: true });
    expect(column?.from.row).toBeNull();

    const [row] = findReferences("=SUM(1:1)");
    expect(row).toMatchObject({ text: "1:1", pair: true });
    expect(row?.from.column).toBeNull();
  });

  it("never reads a lone column or row letter as a reference", () => {
    expect(texts("=SUM(A)")).toEqual([]);
    expect(texts("=SUM(3)")).toEqual([]);
  });

  it("reads a sheet prefix, quoted or not", () => {
    expect(findReferences("=Sheet2!A1")[0]).toMatchObject({
      text: "Sheet2!A1",
      prefix: "Sheet2!",
      sheet: "Sheet2",
    });
    expect(findReferences("='P&L 2025'!A1:B2")[0]).toMatchObject({
      text: "'P&L 2025'!A1:B2",
      prefix: "'P&L 2025'!",
      sheet: "P&L 2025",
      pair: true,
    });
  });

  it("reads an unquoted sheet name with accented letters", () => {
    expect(findReferences("=Pārskats!A1")[0]).toMatchObject({
      text: "Pārskats!A1",
      sheet: "Pārskats",
    });
  });

  it("keeps a ! and a doubled apostrophe inside a quoted sheet name", () => {
    expect(findReferences("='Sheet!1'!A1")[0]).toMatchObject({
      text: "'Sheet!1'!A1",
      sheet: "Sheet!1",
    });
    expect(findReferences("='Bob''s'!A1")[0]).toMatchObject({
      text: "'Bob''s'!A1",
      sheet: "Bob's",
    });
  });

  it("reads another workbook's sheet as its own name", () => {
    expect(findReferences("=[Book1.xlsx]Sheet1!A1")[0]).toMatchObject({
      text: "[Book1.xlsx]Sheet1!A1",
      sheet: "[Book1.xlsx]Sheet1",
    });
  });

  it("ignores text inside double quotes", () => {
    expect(texts('="A1"&A1')).toEqual(["A1"]);
    expect(texts('="say ""A1"" now"')).toEqual([]);
  });

  it("ignores a structured reference", () => {
    expect(texts("=SUM(Table1[Col])")).toEqual([]);
    expect(texts("=SUM(Table1[[#Headers],[A1]])")).toEqual([]);
  });

  it("ignores a function name that reads like a cell", () => {
    expect(texts("=LOG10(A1)")).toEqual(["A1"]);
  });

  it("ignores a defined name that ends in a cell address", () => {
    expect(texts("=Tax_A1")).toEqual([]);
    expect(texts("=A1x")).toEqual([]);
    expect(texts("=XA1")).toEqual(["XA1"]);
  });

  it("refuses an address past the grid", () => {
    expect(texts("=XFE1")).toEqual([]);
    expect(texts("=A1048577")).toEqual([]);
    expect(texts("=XFD1048576")).toEqual(["XFD1048576"]);
  });

  it("leaves a three-dimensional reference to its sheet prefix", () => {
    expect(findReferences("=SUM(Sheet1:Sheet3!A1)")[0]).toMatchObject({
      sheet: "Sheet3",
      text: "Sheet3!A1",
    });
  });
});

describe("rewriteReferences", () => {
  it("splices replacements without touching the rest", () => {
    const out = rewriteReferences("=A1+SUM(B2:C3)+1", (ref) =>
      ref.pair ? "D4:E5" : null,
    );
    expect(out).toBe("=A1+SUM(D4:E5)+1");
  });

  it("returns the formula unchanged when nothing is rewritten", () => {
    expect(rewriteReferences("=A1+B2", () => null)).toBe("=A1+B2");
  });
});
