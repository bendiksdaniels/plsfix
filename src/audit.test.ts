import { describe, expect, it } from "vitest";
import { auditGrid } from "./audit";

describe("cells with nothing to compare", () => {
  it("leaves constants, text and blanks unmarked", () => {
    expect(auditGrid([[1, "Revenue", null, ""]])).toEqual([
      ["none", "none", "none", "none"],
    ]);
  });

  it("leaves a standalone formula unmarked", () => {
    expect(
      auditGrid([
        [null, 12, null],
        [3, "=SUM(R[-1]C:R[-1]C)", "Total"],
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
