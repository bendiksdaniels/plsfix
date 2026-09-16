import { describe, expect, it } from "vitest";
import { auditGrid, consistentRegion } from "./audit";

describe("cells with nothing to compare", () => {
  it("leaves constants, text and blanks unmarked", () => {
    expect(auditGrid([[1, "Revenue", null, ""]])).toEqual([
      ["none", "none", "none", "none"],
    ]);
  });

  it("leaves a standalone formula unmarked", () => {
    // The 3 that used to sit left of the formula is null here instead: a
    // number beside a formula is now "typed" (see the describe block below),
    // and this fixture's job is the formula's own "none", not that one.
    expect(
      auditGrid([
        [null, 12, null],
        [null, "=SUM(R[-1]C:R[-1]C)", "Total"],
        [null, null, null],
      ]),
    ).toEqual([
      ["none", "none", "none"],
      ["none", "none", "none"],
      ["none", "none", "none"],
    ]);
  });

  it("returns an empty grid for an empty selection", () => {
    expect(auditGrid([])).toEqual([]);
  });
});

describe("consistent formulas", () => {
  it("marks a filled-right row horizontal", () => {
    const cell = "=RC[-1]*1.05";
    expect(auditGrid([[cell, cell, cell, cell]])).toEqual([
      ["horizontal", "horizontal", "horizontal", "horizontal"],
    ]);
  });

  it("marks a filled-down column vertical", () => {
    const cell = "=R[-1]C+RC[-1]";
    expect(auditGrid([[cell], [cell], [cell]])).toEqual([
      ["vertical"],
      ["vertical"],
      ["vertical"],
    ]);
  });

  it("marks a filled block both ways", () => {
    const cell = "=R[-1]C*RC[-1]";
    expect(
      auditGrid([
        [cell, cell],
        [cell, cell],
        [cell, cell],
      ]),
    ).toEqual([
      ["both", "both"],
      ["both", "both"],
      ["both", "both"],
    ]);
  });

  it("marks vertical when only the column matches", () => {
    const down = "=R[-1]C";
    const across = "=RC[-1]";
    expect(
      auditGrid([
        [down, across],
        [down, across],
      ]),
    ).toEqual([
      ["vertical", "vertical"],
      ["vertical", "vertical"],
    ]);
  });
});

describe("deviations", () => {
  it("marks the odd formula mid-row lone", () => {
    const cell = "=RC[-1]*1.05";
    expect(auditGrid([[cell, cell, "=RC[-1]*1.5", cell, cell]])).toEqual([
      ["horizontal", "horizontal", "lone", "horizontal", "horizontal"],
    ]);
  });

  it("marks a deviation at the end of a row lone", () => {
    const cell = "=SUM(R[-3]C:R[-1]C)";
    expect(auditGrid([[cell, cell, "=R[-1]C+42"]])).toEqual([
      ["horizontal", "horizontal", "lone"],
    ]);
  });

  it("marks a deviation inside a block and downgrades its neighbours", () => {
    const cell = "=R[-1]C*RC[-1]";
    expect(
      auditGrid([
        [cell, cell, cell],
        [cell, "=R[-1]C*2", cell],
        [cell, cell, cell],
      ]),
    ).toEqual([
      ["both", "horizontal", "both"],
      ["vertical", "lone", "vertical"],
      ["both", "horizontal", "both"],
    ]);
  });

  it("marks two adjacent mismatched formulas lone", () => {
    expect(auditGrid([["=RC[-1]", "=RC[-2]"]])).toEqual([["lone", "lone"]]);
  });
});

describe("typed numbers inside a formula row", () => {
  it("marks a hardcode typed between two equal formulas, and clears their lone flags", () => {
    const cell = "=RC[-1]*1.05";
    expect(auditGrid([[cell, 0.21, cell]])).toEqual([
      ["horizontal", "typed", "horizontal"],
    ]);
  });

  it("marks a typed number at the row end with one formula neighbour", () => {
    const cell = "=RC[-1]*1.05";
    expect(auditGrid([[42, cell, cell]])).toEqual([
      ["typed", "horizontal", "horizontal"],
    ]);
  });

  it("leaves a number beside text none", () => {
    expect(auditGrid([[5, "Label"]])).toEqual([["none", "none"]]);
  });

  it("leaves a number none when its two neighbours are different formulas", () => {
    const left = "=RC[-1]*1.05";
    const right = "=RC[-1]*1.5";
    expect(auditGrid([[left, 7, right]])).toEqual([["lone", "none", "lone"]]);
  });

  it("stops the formula beside a typed number from reading as lone", () => {
    // G19-style row: the same formula either side of a hardcode, plus an
    // unrelated formula above the last column so the old rule (immediate
    // across neighbours only) would have called it "lone" rather than "none".
    const rowFormula = "=RC[-1]*1.05";
    const otherFormula = "=RC[-1]+9";
    const grid = [
      [null, null, null, otherFormula],
      [rowFormula, rowFormula, 0.21, rowFormula],
      [null, null, null, null],
    ];
    expect(auditGrid(grid)[1]).toEqual([
      "horizontal",
      "horizontal",
      "typed",
      "horizontal",
    ]);
  });

  it("leaves a whole numeric row none between formula rows", () => {
    const cell = "=R[-1]C*1.05";
    const grid = [
      [cell, cell, cell],
      [1, 2, 3],
      [cell, cell, cell],
    ];
    expect(auditGrid(grid)).toEqual([
      ["horizontal", "horizontal", "horizontal"],
      ["none", "none", "none"],
      ["horizontal", "horizontal", "horizontal"],
    ]);
  });
});

// The rectangle "Select consistent region" hands back to Excel: the same R1C1
// equality the overlay classifies with, grown instead of classified.
describe("consistentRegion", () => {
  const cell = "=RC[-1]*1.05";
  const other = "=RC[-1]*1.5";

  it("answers nothing for a cell that holds no formula", () => {
    expect(consistentRegion([[1, "Revenue", null]], 0, 1)).toBeNull();
  });

  it("answers nothing for a cell outside the grid", () => {
    expect(consistentRegion([[cell]], 4, 9)).toBeNull();
    expect(consistentRegion([], 0, 0)).toBeNull();
  });

  it("answers the cell itself when nothing around it matches", () => {
    expect(
      consistentRegion(
        [
          [null, other, null],
          [12, cell, "Total"],
          [null, null, null],
        ],
        1,
        1,
      ),
    ).toEqual({ row: 1, column: 1, rowCount: 1, columnCount: 1 });
  });

  it("grows across a filled-right row from any cell in it", () => {
    expect(consistentRegion([[cell, cell, cell, cell]], 0, 2)).toEqual({
      row: 0,
      column: 0,
      rowCount: 1,
      columnCount: 4,
    });
  });

  it("grows down a filled column", () => {
    expect(consistentRegion([[cell], [cell], [cell]], 2, 0)).toEqual({
      row: 0,
      column: 0,
      rowCount: 3,
      columnCount: 1,
    });
  });

  it("grows to the whole block when every row and column matches", () => {
    const grid = [
      ["Label", null, null, null],
      [null, cell, cell, cell],
      [null, cell, cell, cell],
    ];
    expect(consistentRegion(grid, 1, 2)).toEqual({
      row: 1,
      column: 1,
      rowCount: 2,
      columnCount: 3,
    });
  });

  it("stops at a blank and at a different formula", () => {
    expect(consistentRegion([[other, cell, cell, null, cell]], 0, 1)).toEqual({
      row: 0,
      column: 1,
      rowCount: 1,
      columnCount: 2,
    });
  });

  it("stops a direction whose next row is only partly the same formula", () => {
    const grid = [
      [cell, cell, cell],
      [cell, other, cell],
      [cell, cell, cell],
    ];
    expect(consistentRegion(grid, 0, 0)).toEqual({
      row: 0,
      column: 0,
      rowCount: 1,
      columnCount: 3,
    });
  });

  it("treats a short row as blank rather than reading past its end", () => {
    expect(consistentRegion([[cell, cell], [cell]], 0, 0)).toEqual({
      row: 0,
      column: 0,
      rowCount: 1,
      columnCount: 2,
    });
  });

  // Growth is horizontal first, then vertical, repeated until the rectangle
  // stops changing: an L of matching cells becomes the row, not the column.
  it("takes the row before the column when both would fit an L", () => {
    const grid = [
      [cell, cell],
      [cell, null],
    ];
    expect(consistentRegion(grid, 0, 0)).toEqual({
      row: 0,
      column: 0,
      rowCount: 1,
      columnCount: 2,
    });
  });
});
