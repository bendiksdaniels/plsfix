// Pass-1 hunt: classifyCell's workbook-reference detection against sheet names
// that need Excel's own single-quoting (spaces, apostrophes, and everything
// else that is not a plain identifier). SHEET_CHAR only ever needs to guess at
// the unquoted form: once a reference is quoted, Excel has already resolved
// the ambiguity, and the scan should trust the quote, not a character list.
import { describe, expect, it } from "vitest";
import { classifyCell } from "../../src/classify";

describe("a quoted external reference whose sheet name needs no escaping", () => {
  it("still reads a plain quoted sheet name as external", () => {
    expect(classifyCell("='[Budget.xlsx]Model plan'!$B$4", 4)).toBe("external");
  });
});

describe("a quoted external reference to a sheet name with a hyphen", () => {
  it("reads a dated sheet name as external, not cross-sheet", () => {
    // "Q1-2026" is an ordinary, common sheet name; Excel quotes the whole
    // [Workbook]Sheet reference because of it, but never uses a hyphen as a
    // formula operator inside a quoted reference.
    expect(classifyCell("='[Budget.xlsx]Q1-2026'!A1", 1)).toBe("external");
  });

  it("still reads the unquoted form the same reference would take without one", () => {
    // A bare structured reference followed by a subtraction and an unrelated
    // cross-sheet operand must not be swept up by a widened hyphen allowance:
    // this is the reason the unquoted scan stays narrow.
    expect(classifyCell("=[Amount]-Sheet2!B1", 1)).toBe("crossSheet");
  });
});

describe("a quoted external reference to a sheet name with an ampersand", () => {
  it("reads it as external", () => {
    expect(classifyCell("='[Model.xlsx]R&D'!A1", 1)).toBe("external");
  });
});

describe("a quoted external reference to a sheet name with a doubled apostrophe", () => {
  it("reads it as external past the escaped quote", () => {
    expect(classifyCell("='[Budget.xlsx]O''Brien''s Model'!$B$4", 4)).toBe(
      "external",
    );
  });
});

describe("a quoted external reference to a sheet name in Latvian", () => {
  it("reads unicode letters in the sheet name as external", () => {
    expect(classifyCell("='[Model.xlsx]Piegādātāji'!A1", 1)).toBe("external");
  });

  it("ranks it external above the hardcode it carries", () => {
    expect(classifyCell("='[Model.xlsx]Piegādātāji'!A1*1.05", 1)).toBe(
      "external",
    );
  });
});

describe("the workbook-index form with a sheet name that needs quoting", () => {
  it("still reads a closed workbook's own indexed link as external", () => {
    expect(classifyCell("='[1]Q1-2026'!A1", 1)).toBe("external");
  });
});
