import { describe, expect, it } from "vitest";
import {
  buildCagrFormula,
  detectFillExtent,
  flipSign,
  stepDecimals,
  toggleIfError,
} from "./paste";

describe("detectFillExtent", () => {
  it("sizes the fill from a single neighbour line", () => {
    expect(detectFillExtent([[1, 2, 3, 4, 5]])).toBe(5);
  });

  it("takes the longest of several neighbour lines", () => {
    expect(
      detectFillExtent([
        ["a", "b", "c"],
        [1, 2, 3, 4, 5, 6, 7],
      ]),
    ).toBe(7);
  });

  it("returns zero when every neighbour starts blank", () => {
    expect(
      detectFillExtent([
        [null, 1, 2],
        ["", "x"],
      ]),
    ).toBe(0);
    expect(detectFillExtent([])).toBe(0);
  });

  it("stops at the first gap in a line", () => {
    expect(detectFillExtent([[1, 2, 3, 4, null, 6, 7]])).toBe(4);
    expect(detectFillExtent([[1, 2, 3, 4, "", 6]])).toBe(4);
  });
});

describe("flipSign", () => {
  it("negates numbers and leaves text alone", () => {
    expect(flipSign([[10, -4, "Text", null]])).toEqual([
      [-10, 4, "Text", null],
    ]);
  });

  it("wraps a formula and unwraps it again", () => {
    expect(flipSign([["=A1+B1"]])).toEqual([["=-(A1+B1)"]]);
    expect(flipSign([["=-(A1+B1)"]])).toEqual([["=A1+B1"]]);
  });

  it("wraps a formula whose leading parenthesis closes early", () => {
    expect(flipSign([["=-(A1)+B1"]])).toEqual([["=-(-(A1)+B1)"]]);
  });
});

describe("stepDecimals", () => {
  it("adds and removes one decimal place", () => {
    expect(stepDecimals("#,##0", 1)).toBe("#,##0.0");
    expect(stepDecimals("#,##0.00", -1)).toBe("#,##0.0");
    expect(stepDecimals("#,##0.0", -1)).toBe("#,##0");
    expect(stepDecimals("0.0%", 1)).toBe("0.00%");
    expect(stepDecimals("0.0%", -1)).toBe("0%");
    expect(stepDecimals("0.00", -1)).toBe("0.0");
    expect(stepDecimals("0.00", 1)).toBe("0.000");
  });

  it("steps every section of the shipped families", () => {
    expect(stepDecimals("#,##0;[Red](#,##0);-", 1)).toBe(
      "#,##0.0;[Red](#,##0.0);-",
    );
    expect(stepDecimals("0.0%;[Red](0.0%);-", 1)).toBe("0.00%;[Red](0.00%);-");
    expect(stepDecimals('0.0"x";[Red](0.0"x");-', 1)).toBe(
      '0.00"x";[Red](0.00"x");-',
    );
    expect(stepDecimals("€ #,##0.0;[Red](€ #,##0.0);-", -1)).toBe(
      "€ #,##0;[Red](€ #,##0);-",
    );
    expect(stepDecimals("[$€-x-euro2] #,##0", 1)).toBe("[$€-x-euro2] #,##0.0");
  });

  it("keeps quoted literals and trailing scale commas intact", () => {
    expect(stepDecimals('0.0"x"', 1)).toBe('0.00"x"');
    expect(stepDecimals('0.0" ; "', 1)).toBe('0.00" ; "');
    expect(stepDecimals("€ #,##0,", 1)).toBe("€ #,##0.0,");
  });

  it("leaves formats without digit placeholders unchanged", () => {
    expect(stepDecimals("General", 1)).toBe("General");
    expect(stepDecimals("General", -1)).toBe("General");
    expect(stepDecimals("@", -1)).toBe("@");
    expect(stepDecimals("dd.mm.yyyy", 1)).toBe("dd.mm.yyyy");
    expect(stepDecimals("#,##0;[Red](#,##0);-", -1)).toBe(
      "#,##0;[Red](#,##0);-",
    );
  });
});

describe("toggleIfError", () => {
  it("wraps unguarded formulas and leaves constants alone", () => {
    expect(toggleIfError([["=A1/B1", 25, "Text", null]], "0")).toEqual([
      ["=IFERROR(A1/B1,0)", 25, "Text", null],
    ]);
  });

  it("unwraps a formula it already guarded", () => {
    expect(toggleIfError([["=IFERROR(A1/B1,0)"]], "0")).toEqual([["=A1/B1"]]);
  });

  it("keeps a nested IFERROR when unwrapping the outer one", () => {
    expect(toggleIfError([["=IFERROR(IFERROR(A1,B1),0)"]], "0")).toEqual([
      ["=IFERROR(A1,B1)"],
    ]);
  });

  it("ignores commas inside nested calls and quoted strings", () => {
    expect(
      toggleIfError([['=IFERROR(VLOOKUP(A1,B:C,2,FALSE),"n,a")']], "0"),
    ).toEqual([["=VLOOKUP(A1,B:C,2,FALSE)"]]);
    expect(toggleIfError([['=IF(A1,"x,y",0)']], '""')).toEqual([
      ['=IFERROR(IF(A1,"x,y",0),"")'],
    ]);
  });

  it("wraps a formula that only starts with IFERROR", () => {
    expect(toggleIfError([["=IFERROR(A1,0)+B1"]], "0")).toEqual([
      ["=IFERROR(IFERROR(A1,0)+B1,0)"],
    ]);
  });
});

describe("buildCagrFormula", () => {
  it("builds the compound growth formula", () => {
    expect(buildCagrFormula("B5", "F5", 4)).toBe("=(F5/B5)^(1/4)-1");
    expect(buildCagrFormula("C2", "C9", 7)).toBe("=(C9/C2)^(1/7)-1");
  });
});
