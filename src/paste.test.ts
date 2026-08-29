import { describe, expect, it } from "vitest";
import {
  absoluteRef,
  buildCagrFormula,
  buildRoundFormula,
  detectFillExtent,
  flipSign,
  formatDecimals,
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

describe("formatDecimals", () => {
  it("counts the decimals a format prints", () => {
    expect(formatDecimals("#,##0")).toBe(0);
    expect(formatDecimals("#,##0.00")).toBe(2);
    expect(formatDecimals("0.000")).toBe(3);
    expect(formatDecimals("#,##0;[Red](#,##0);-")).toBe(0);
    expect(formatDecimals("#,##0.0;[Red](#,##0.0);-")).toBe(1);
    expect(formatDecimals("[$€-x-euro2] #,##0.00")).toBe(2);
  });

  it("follows the scale a format applies before printing", () => {
    // A percentage is stored a hundred times smaller than it prints, and each
    // trailing comma prints a thousand times smaller than it is stored.
    expect(formatDecimals("0.0%")).toBe(3);
    expect(formatDecimals("0.0%;[Red](0.0%);-")).toBe(3);
    expect(formatDecimals("0%")).toBe(2);
    expect(formatDecimals("#,##0,")).toBe(-3);
    expect(formatDecimals("#,##0.0,,")).toBe(-5);
  });

  it("returns null when the format prints no digits", () => {
    expect(formatDecimals("General")).toBeNull();
    expect(formatDecimals("@")).toBeNull();
    expect(formatDecimals("dd.mm.yyyy")).toBeNull();
  });

  it("reads through quoted and bracketed literals", () => {
    expect(formatDecimals('0.0"%"')).toBe(1);
    expect(formatDecimals('#,##0.00" ; "')).toBe(2);
  });
});

describe("absoluteRef", () => {
  it("locks a local reference to its cells", () => {
    expect(absoluteRef("A1:A5")).toBe("$A$1:$A$5");
    expect(absoluteRef("B2:F2")).toBe("$B$2:$F$2");
    expect(absoluteRef("AA10")).toBe("$AA$10");
  });

  it("leaves a sheet prefix and existing anchors alone", () => {
    expect(absoluteRef("Model!A1:A3")).toBe("Model!$A$1:$A$3");
    expect(absoluteRef("'Q3 model'!C4")).toBe("'Q3 model'!$C$4");
    expect(absoluteRef("$A$1:$A$5")).toBe("$A$1:$A$5");
  });
});

describe("buildRoundFormula", () => {
  it("passes the whole group, the position and the precision", () => {
    expect(buildRoundFormula("$A$1:$A$3", 1, 0)).toBe(
      "=SMT.ROUND($A$1:$A$3,1,0)",
    );
    expect(buildRoundFormula("$B$2:$D$2", 3, 2)).toBe(
      "=SMT.ROUND($B$2:$D$2,3,2)",
    );
  });
});
