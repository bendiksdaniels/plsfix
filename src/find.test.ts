import { describe, expect, it } from "vitest";
import {
  cellAddress,
  cellText,
  COMMENT_ROW,
  COMMENT_TEXT_LIMIT,
  FIND_HIT_CAP,
  includesQuery,
  matchCells,
  matchComments,
  rankHits,
  type FindGrid,
} from "./find";

const LOOSE = { matchCase: false, inFormulas: false };
const FORMULAS = { matchCase: false, inFormulas: true };

// values and formulas arrive as two grids of the same shape; a constant cell
// repeats its value in both, which is what Excel hands back.
function grid(rows: (string | number | boolean | null)[][]): FindGrid {
  return { values: rows, formulas: rows };
}

describe("matchCells", () => {
  it("matches a substring in any casing by default", () => {
    expect(matchCells(grid([["Revenue growth"]]), "venue", LOOSE)).toEqual([
      { row: 0, col: 0, text: "Revenue growth" },
    ]);
    expect(matchCells(grid([["Revenue growth"]]), "REVENUE", LOOSE)).toEqual([
      { row: 0, col: 0, text: "Revenue growth" },
    ]);
  });

  it("respects match case when asked", () => {
    const options = { matchCase: true, inFormulas: false };
    expect(matchCells(grid([["Revenue"]]), "revenue", options)).toEqual([]);
    expect(matchCells(grid([["Revenue"]]), "Revenue", options)).toEqual([
      { row: 0, col: 0, text: "Revenue" },
    ]);
  });

  it("reads numbers and booleans the way Excel shows them", () => {
    expect(matchCells(grid([[1234.5, true]]), "34.5", LOOSE)).toEqual([
      { row: 0, col: 0, text: "1234.5" },
    ]);
    expect(matchCells(grid([[1234.5, true]]), "TRUE", LOOSE)).toEqual([
      { row: 0, col: 1, text: "TRUE" },
    ]);
  });

  it("never matches a blank cell, whatever the query", () => {
    expect(matchCells(grid([[null, ""]]), "", LOOSE)).toEqual([]);
    expect(matchCells(grid([[null, ""]]), "a", LOOSE)).toEqual([]);
  });

  it("leaves formulas alone unless the box is ticked", () => {
    const sheet: FindGrid = {
      values: [[42]],
      formulas: [["=Drivers!B2*2"]],
    };
    expect(matchCells(sheet, "Drivers", LOOSE)).toEqual([]);
    expect(matchCells(sheet, "Drivers", FORMULAS)).toEqual([
      { row: 0, col: 0, text: "=Drivers!B2*2" },
    ]);
  });

  it("shows the value when the value is what matched", () => {
    const sheet: FindGrid = {
      values: [["Drivers copy"]],
      formulas: [["=Drivers!B2"]],
    };
    expect(matchCells(sheet, "Drivers", FORMULAS)).toEqual([
      { row: 0, col: 0, text: "Drivers copy" },
    ]);
  });

  it("counts a constant once, not once per grid", () => {
    const sheet: FindGrid = { values: [["Total"]], formulas: [["Total"]] };
    expect(matchCells(sheet, "Total", FORMULAS)).toHaveLength(1);
  });

  it("reads down and across, in that order", () => {
    const sheet = grid([
      ["x", "x"],
      ["x", "x"],
    ]);
    expect(
      matchCells(sheet, "x", LOOSE).map((hit) => [hit.row, hit.col]),
    ).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
  });

  it("stops at the cap however many cells match", () => {
    const rows = Array.from({ length: 30 }, () =>
      Array.from({ length: 30 }, () => "hit"),
    );
    expect(matchCells(grid(rows), "hit", LOOSE)).toHaveLength(FIND_HIT_CAP);
  });

  it("tolerates a ragged grid", () => {
    const sheet: FindGrid = { values: [["a"], ["a", "a"]], formulas: [[]] };
    expect(matchCells(sheet, "a", FORMULAS)).toHaveLength(3);
  });
});

describe("rankHits", () => {
  it("orders by sheet, then row, then column", () => {
    const hits = [
      { sheetIndex: 1, row: 4, col: 0 },
      { sheetIndex: 0, row: 2, col: 3 },
      { sheetIndex: 0, row: 2, col: 1 },
      { sheetIndex: -1, row: 0, col: 0 },
      { sheetIndex: 0, row: -1, col: -1 },
    ];
    expect(rankHits(hits)).toEqual([
      { sheetIndex: -1, row: 0, col: 0 },
      { sheetIndex: 0, row: -1, col: -1 },
      { sheetIndex: 0, row: 2, col: 1 },
      { sheetIndex: 0, row: 2, col: 3 },
      { sheetIndex: 1, row: 4, col: 0 },
    ]);
  });

  it("leaves the list it was given alone", () => {
    const hits = [
      { sheetIndex: 1, row: 0, col: 0 },
      { sheetIndex: 0, row: 0, col: 0 },
    ];
    rankHits(hits);
    expect(hits[0]?.sheetIndex).toBe(1);
  });

  it("keeps the first hits when the workbook is over the cap", () => {
    const hits = Array.from({ length: FIND_HIT_CAP + 25 }, (_, index) => ({
      sheetIndex: 0,
      row: FIND_HIT_CAP + 25 - index,
      col: 0,
    }));
    const ranked = rankHits(hits);
    expect(ranked).toHaveLength(FIND_HIT_CAP);
    expect(ranked[0]?.row).toBe(1);
  });
});

describe("matchComments", () => {
  const thread = (content: string, author = "Anna Ozola") => ({
    sheet: "Model",
    address: "B2",
    content,
    author,
  });

  it("matches the content in any casing and says where it sat", () => {
    const comments = [thread("Ignore"), thread("Check the WACC here")];
    expect(matchComments(comments, "wacc", LOOSE)).toEqual([
      { order: 1, text: "Check the WACC here" },
    ]);
  });

  it("matches the author as well as what they wrote", () => {
    expect(matchComments([thread("Looks fine")], "ozola", LOOSE)).toEqual([
      { order: 0, text: "Looks fine" },
    ]);
  });

  it("respects match case on both content and author", () => {
    const options = { matchCase: true, inFormulas: false };
    expect(matchComments([thread("Revenue")], "revenue", options)).toEqual([]);
    expect(matchComments([thread("Revenue")], "Revenue", options)).toHaveLength(
      1,
    );
    expect(matchComments([thread("x")], "Anna", options)).toHaveLength(1);
  });

  it("shows the opening of a long thread, not the whole of it", () => {
    const long = "note ".repeat(60);
    const [hit] = matchComments([thread(long)], "note", LOOSE);
    expect(hit?.text).toHaveLength(COMMENT_TEXT_LIMIT);
    expect(hit?.text).toBe(long.slice(0, COMMENT_TEXT_LIMIT));
  });

  it("never matches on an empty query", () => {
    expect(matchComments([thread("anything")], "", LOOSE)).toEqual([]);
  });

  it("stops at the cap however many comments match", () => {
    const comments = Array.from({ length: FIND_HIT_CAP + 40 }, () =>
      thread("hit"),
    );
    expect(matchComments(comments, "hit", LOOSE)).toHaveLength(FIND_HIT_CAP);
  });
});

describe("rankHits over comments", () => {
  it("lists a sheet's comments after its cells, before the next sheet", () => {
    const hits = [
      { sheetIndex: 1, row: 0, col: 0 },
      { sheetIndex: 0, row: COMMENT_ROW, col: 1 },
      { sheetIndex: 0, row: COMMENT_ROW, col: 0 },
      { sheetIndex: 0, row: 900, col: 0 },
    ];
    expect(rankHits(hits)).toEqual([
      { sheetIndex: 0, row: 900, col: 0 },
      { sheetIndex: 0, row: COMMENT_ROW, col: 0 },
      { sheetIndex: 0, row: COMMENT_ROW, col: 1 },
      { sheetIndex: 1, row: 0, col: 0 },
    ]);
  });

  it("spends one cap on cells and comments together", () => {
    const cells = Array.from({ length: FIND_HIT_CAP }, (_unused, row) => ({
      sheetIndex: 0,
      row,
      col: 0,
    }));
    const ranked = rankHits([
      ...cells,
      { sheetIndex: 0, row: COMMENT_ROW, col: 0 },
    ]);

    expect(ranked).toHaveLength(FIND_HIT_CAP);
    expect(ranked.some((hit) => hit.row === COMMENT_ROW)).toBe(false);
  });
});

describe("includesQuery and cellText", () => {
  it("never matches on an empty query", () => {
    expect(includesQuery("anything", "", false)).toBe(false);
  });

  it("reads a blank cell as empty text", () => {
    expect(cellText(null)).toBe("");
    expect(cellText(false)).toBe("FALSE");
    expect(cellText(0)).toBe("0");
  });
});

describe("cellAddress", () => {
  it("names the column the way Excel does", () => {
    expect(cellAddress(0, 0)).toBe("A1");
    expect(cellAddress(6, 1)).toBe("B7");
    expect(cellAddress(0, 25)).toBe("Z1");
    expect(cellAddress(0, 26)).toBe("AA1");
    expect(cellAddress(99, 701)).toBe("ZZ100");
    expect(cellAddress(0, 702)).toBe("AAA1");
  });
});
