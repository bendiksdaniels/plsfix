// Pass-1 hunt: the IFERROR toggle's scanner. A comma or paren inside a quoted
// string, a quoted sheet name or a table bracket is never the guard's own, in
// either direction of the toggle; plus a number format ending mid-escape.
import { describe, expect, it } from "vitest";
import {
  countIfErrorToggle,
  flipSign,
  formatDecimals,
  stepDecimals,
  toggleIfError,
} from "../../src/paste";

describe("a quoted comma inside the guarded expression, not the fallback", () => {
  it("does not mistake the comma inside the string for the guard's own separator", () => {
    // Every existing quoted-string case puts the comma in the FALLBACK, after
    // the top-level separator is already found; this one puts it before, so
    // the scanner must open and close a quote while still looking for that
    // separator (src/paste.ts unwrapIfError, the quoted-closing branch).
    expect(toggleIfError([['=IFERROR(A1&"a,b",0)']], "0")).toEqual([
      ['=A1&"a,b"'],
    ]);
  });

  it("re-guards the same shape identically, round trip", () => {
    const once = toggleIfError([['=IFERROR(A1&"a,b",0)']], "0");
    expect(toggleIfError(once, "0")).toEqual([['=IFERROR(A1&"a,b",0)']]);
  });
});

describe("a doubled quote (Excel's escaped literal quote) inside that expression", () => {
  it("still finds the real top-level comma past two escaped quote pairs", () => {
    // The literal text is: x"y,z"w - two escaped quotes, with the guard's own
    // comma right after the string closes for real. A parser that does not
    // know about doubling would still get this right here, because "" is two
    // adjacent characters with no room for a stray comma between them - but
    // it is exactly the kind of thing worth pinning by name.
    const formula = '=IFERROR(A1&"x""y,z""w",0)';
    expect(toggleIfError([[formula]], "0")).toEqual([['=A1&"x""y,z""w"']]);
  });

  it("keeps behaving once the guarded expression's string comes first", () => {
    const formula = '=IFERROR("a,b"&A1,"n/a")';
    expect(toggleIfError([[formula]], "0")).toEqual([['="a,b"&A1']]);
  });
});

describe("a quoted sheet name or a table bracket inside the guarded expression", () => {
  // Excel refuses the half formulas the old scanner wrote ("='Sales"), and the
  // receipt still said "stripped 1": the scanner only knew "..." strings.
  it.each([
    ["=IFERROR('Sales, EU'!A1/B1,0)", "='Sales, EU'!A1/B1"],
    [
      "=IFERROR(Table1[[#This Row],[Price]]/B1,0)",
      "=Table1[[#This Row],[Price]]/B1",
    ],
    ["=IFERROR('Q1 (EU)'!A1,0)", "='Q1 (EU)'!A1"],
    ["=IFERROR('O''Brien, Ltd'!A1,0)", "='O''Brien, Ltd'!A1"],
    ["=IFERROR(T[[#This Row],[a']b]]/2,0)", "=T[[#This Row],[a']b]]/2"],
  ])("unwraps %s to the whole guarded expression", (guarded, plain) => {
    expect(toggleIfError([[guarded]], "0")).toEqual([[plain]]);
    expect(countIfErrorToggle([[guarded]])).toEqual({ added: 0, stripped: 1 });
  });

  it("round-trips a quoted sheet name with a comma: guard, then unwrap", () => {
    const guarded = toggleIfError([["='Sales, EU'!A1/B1"]], "0");
    expect(guarded).toEqual([["=IFERROR('Sales, EU'!A1/B1,0)"]]);
    expect(toggleIfError(guarded, "0")).toEqual([["='Sales, EU'!A1/B1"]]);
  });

  it("flips the sign back over a sheet name holding parentheses", () => {
    expect(flipSign([["=-('Q1 (EU)'!A1)"]])).toEqual([["='Q1 (EU)'!A1"]]);
  });
});

describe("a number format that ends on a lone trailing backslash", () => {
  it("does not read past the end of the string", () => {
    // scanLiterals marks the backslash itself literal and has nothing left
    // to also mark as its escaped partner; no crash, no digits stepped past
    // where the format actually ends.
    expect(() => stepDecimals("0.0\\", 1)).not.toThrow();
    expect(stepDecimals("0.0\\", 1)).toBe("0.00\\");
    expect(formatDecimals("0.0\\")).toBe(1);
  });
});
