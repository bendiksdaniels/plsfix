import { describe, expect, it } from "vitest";
import { unpivot } from "./reshape";

describe("unpivot", () => {
  it("turns a wide block into one line per populated cell", () => {
    expect(
      unpivot([
        ["", "North", "South"],
        ["Q1", 10, 20],
        ["Q2", 30, 40],
      ]),
    ).toEqual([
      ["Q1", "North", 10],
      ["Q1", "South", 20],
      ["Q2", "North", 30],
      ["Q2", "South", 40],
    ]);
  });

  it("keeps numbers, text and booleans as they are", () => {
    expect(
      unpivot([
        ["", "Value", "Note", "Flag"],
        ["Row", -1.5, "check", true],
      ]),
    ).toEqual([
      ["Row", "Value", -1.5],
      ["Row", "Note", "check"],
      ["Row", "Flag", true],
    ]);
  });

  it("skips blank cells rather than writing empty rows", () => {
    expect(
      unpivot([
        ["", "North", "South"],
        ["Q1", null, 20],
        ["Q2", "", ""],
      ]),
    ).toEqual([["Q1", "South", 20]]);
  });

  it("skips a column with no header and a row with no key", () => {
    expect(
      unpivot([
        ["", "North", "", "South"],
        ["Q1", 10, 99, 20],
        ["", 11, 98, 21],
        [null, 12, 97, 22],
      ]),
    ).toEqual([
      ["Q1", "North", 10],
      ["Q1", "South", 20],
    ]);
  });

  it("reads a zero as a value and a blank string as a gap", () => {
    expect(
      unpivot([
        ["", "A"],
        ["Row", 0],
      ]),
    ).toEqual([["Row", "A", 0]]);
  });

  it("returns nothing when every value cell is blank", () => {
    expect(
      unpivot([
        ["", "North"],
        ["Q1", ""],
      ]),
    ).toEqual([]);
  });

  it("refuses a grid without a header row and a key column", () => {
    const message =
      "unpivot: need a header row, a key column and one column of values";
    expect(() => unpivot([])).toThrow(message);
    expect(() => unpivot([["", "North"]])).toThrow(message);
    expect(() => unpivot([["Q1"], ["Q2"]])).toThrow(message);
  });
});
