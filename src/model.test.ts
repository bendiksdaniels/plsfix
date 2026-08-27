import { describe, expect, it } from "vitest";
import {
  analyzeGrid,
  makeFormatGrid,
  scaleCells,
  wrapFormulasWithIfError,
} from "./model";

describe("model helpers", () => {
  it("summarizes formulas, errors and blank cells", () => {
    expect(
      analyzeGrid(
        [["=A1+B1", 10], ["", "=1/0"]],
        [[15, 10], ["", "#DIV/0!"]],
      ),
    ).toEqual({ cells: 4, formulas: 2, errors: 1, blanks: 1 });
  });

  it("wraps formulas once and preserves constants", () => {
    expect(
      wrapFormulasWithIfError([["=A1/B1", "=IFERROR(C1,0)", 25, "Text"]]),
    ).toEqual([["=IFERROR(A1/B1,0)", "=IFERROR(C1,0)", 25, "Text"]]);
  });

  it("scales numeric constants and formulas", () => {
    expect(scaleCells([[10, "=A1+B1", "Text"]], 1000)).toEqual([
      [10_000, "=(A1+B1)*1000", "Text"],
    ]);
    expect(scaleCells([[10_000, "=A1+B1"]], 0.001)).toEqual([
      [10, "=(A1+B1)/1000"],
    ]);
  });

  it("creates a number-format matrix matching the selection", () => {
    expect(makeFormatGrid(2, 3, "0.0%")).toEqual([
      ["0.0%", "0.0%", "0.0%"],
      ["0.0%", "0.0%", "0.0%"],
    ]);
  });
});
