// The comps block, pure: which row is a header, which columns hold numbers, and
// the six live formulas each numeric column gets under the table.

import { describe, expect, it } from "vitest";
import {
  isHeaderRow,
  readCompsBlock,
  STATS_GAP_ROWS,
  STATS_LABELS,
  statsGrid,
  statsPresets,
} from "./stats";
import { type CellValue } from "./model";

const HEADED: CellValue[][] = [
  ["Company", "EV/EBITDA", "Note", "P/E"],
  ["Alpha", 8.1, "listed", 12],
  ["Beta", 9.4, "listed", 14],
  ["Gamma", 7.2, "private", 11],
];

describe("isHeaderRow", () => {
  it("calls a row of column titles a header", () => {
    expect(isHeaderRow(["Company", "EV/EBITDA", "P/E"])).toBe(true);
  });

  it("allows the blank corner a comps table usually has", () => {
    expect(isHeaderRow([null, "EV/EBITDA", "P/E"])).toBe(true);
  });

  it("is not a header when any cell holds a number", () => {
    expect(isHeaderRow(["Alpha", 8.1, 12])).toBe(false);
  });

  it("is not a header when the row holds no text at all", () => {
    expect(isHeaderRow([null, "", null])).toBe(false);
  });
});

describe("readCompsBlock", () => {
  it("drops the header row and keeps the numeric columns", () => {
    const block = readCompsBlock(HEADED, 4, 1);
    expect(block).toMatchObject({
      headed: true,
      firstDataRow: 5,
      lastDataRow: 7,
      firstColumn: 1,
      numericColumns: [1, 3],
    });
  });

  it("reads a header-free table as data rows only", () => {
    const block = readCompsBlock(HEADED.slice(1), 0, 0);
    expect(block).toMatchObject({
      headed: false,
      firstDataRow: 0,
      lastDataRow: 2,
      numericColumns: [1, 3],
    });
  });

  it("refuses fewer than two data rows", () => {
    expect(() => readCompsBlock(HEADED.slice(0, 2), 0, 0)).toThrow(
      "Comps stats need a table with at least two data rows.",
    );
  });

  it("refuses a table with no numeric column beside the labels", () => {
    expect(() =>
      readCompsBlock(
        [
          ["Company", "Note"],
          ["Alpha", "n/a"],
          ["Beta", "n/a"],
        ],
        0,
        0,
      ),
    ).toThrow("Comps stats need at least one column of numbers.");
  });

  it("never treats the label column as numeric, however it is filled", () => {
    expect(() =>
      readCompsBlock(
        [
          [1, "x"],
          [2, "y"],
        ],
        0,
        0,
      ),
    ).toThrow("Comps stats need at least one column of numbers.");
  });
});

describe("statsGrid", () => {
  it("writes six labelled rows of live formulas over the data span", () => {
    const block = readCompsBlock(HEADED, 4, 1);
    expect(statsGrid(block, 4)).toEqual([
      ["Min", "=MIN(C6:C8)", "", "=MIN(E6:E8)"],
      [
        "25th percentile",
        "=PERCENTILE.INC(C6:C8,0.25)",
        "",
        "=PERCENTILE.INC(E6:E8,0.25)",
      ],
      ["Median", "=MEDIAN(C6:C8)", "", "=MEDIAN(E6:E8)"],
      ["Mean", "=AVERAGE(C6:C8)", "", "=AVERAGE(E6:E8)"],
      [
        "75th percentile",
        "=PERCENTILE.INC(C6:C8,0.75)",
        "",
        "=PERCENTILE.INC(E6:E8,0.75)",
      ],
      ["Max", "=MAX(C6:C8)", "", "=MAX(E6:E8)"],
    ]);
  });

  it("labels the rows in the order a comps page reads", () => {
    expect([...STATS_LABELS]).toEqual([
      "Min",
      "25th percentile",
      "Median",
      "Mean",
      "75th percentile",
      "Max",
    ]);
  });

  it("keeps one blank row between the table and the block", () => {
    expect(STATS_GAP_ROWS).toBe(1);
  });

  it("spans a column past Z the way Excel names it", () => {
    const block = readCompsBlock(
      [
        ["Alpha", 1],
        ["Beta", 2],
      ],
      0,
      25,
    );
    expect(statsGrid(block, 2)[0]).toEqual(["Min", "=MIN(AA1:AA2)"]);
  });
});

describe("statsPresets", () => {
  it("gives the labels the plain look and the numbers the formula look", () => {
    const block = readCompsBlock(HEADED, 4, 1);
    expect(statsPresets(block, 4)[0]).toEqual([
      "label",
      "formula",
      null,
      "formula",
    ]);
  });
});
