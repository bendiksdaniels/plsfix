// Pass-2 property: every A1 reference formatReference writes has to come back
// out of findReferences exactly as written - the round trip the whole paste
// suite depends on to know which cells a formula points at. Hammered with
// fast-check over generated corners, $ markers, whole rows/columns and every
// shape of sheet prefix (plain, quoted, apostrophes, unicode, [Book]Sheet!).
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  findReferences,
  formatPart,
  formatReference,
  GRID_COLUMNS,
  GRID_ROWS,
  type RefPart,
  sheetPrefix,
} from "../../src/formula-refs";

const SEED = 20260927;
const RUNS = 2000;

describe("an astral-plane sheet name (regression, found by the round-trip property)", () => {
  // U+1D400 MATHEMATICAL BOLD CAPITAL A: a real Unicode letter (\p{L}) made of
  // a surrogate pair. sheetPrefix's whole-string PLAIN_SHEET test sees one
  // letter and writes it bare; readPrefix's unquoted scanner (src/formula-refs.ts
  // readPrefix, the `else` branch) walks the SAME text one UTF-16 code unit at
  // a time, and a lone surrogate half never matches \p{L} on its own.
  const ASTRAL = "\u{1D400}";

  it("writes the astral sheet name bare", () => {
    expect(sheetPrefix(ASTRAL)).toBe(`${ASTRAL}!`);
  });

  it("reads its own bare prefix back, sheet-qualified, not as an unqualified A1", () => {
    const formula = `=${sheetPrefix(ASTRAL)}A1`;
    const [ref] = findReferences(formula);
    expect(ref?.sheet).toBe(ASTRAL);
    expect(ref?.text).toBe(`${sheetPrefix(ASTRAL)}A1`);
  });

  it("does not let an astral character in a defined name fuse with the address after it", () => {
    // The plain-ASCII analogue already pinned in formula-refs.test.ts: a name
    // ending in a cell-shaped tail with no separator is not a reference at
    // all ("=XA1+1" finds nothing). An astral leading character must refuse
    // the same "A1" for the same reason - it is the tail of one longer name.
    const formula = `=${ASTRAL}A1+1`;
    expect(findReferences(formula)).toEqual([]);
  });

  it("keeps the rest of a longer name once an astral character sits inside it", () => {
    const name = `Model${ASTRAL}2`;
    const formula = `=${sheetPrefix(name)}A1`;
    const [ref] = findReferences(formula);
    expect(ref?.sheet).toBe(name);
  });
});

// $A$1, A1, whole rows/columns and ranges, every $ combination.
const cornerArb: fc.Arbitrary<{
  column: number | null;
  row: number | null;
  columnAbsolute: boolean;
  rowAbsolute: boolean;
}> = fc.record({
  column: fc.integer({ min: 0, max: GRID_COLUMNS - 1 }),
  row: fc.integer({ min: 0, max: GRID_ROWS - 1 }),
  columnAbsolute: fc.boolean(),
  rowAbsolute: fc.boolean(),
});

// A full reference shape: a single cell, or a pair that is a same-kind range
// (two cells, two whole columns or two whole rows - readReference's own
// sameKind rule, so every generated pair is one the scanner can recognize).
const refShapeArb = fc.oneof(
  cornerArb.map((corner) => ({
    pair: false as const,
    from: corner,
    to: corner,
  })),
  fc.tuple(cornerArb, cornerArb).map(([from, to]) => ({
    pair: true as const,
    from,
    to,
  })),
  fc.tuple(cornerArb, cornerArb).map(([from, to]) => ({
    pair: true as const,
    from: { ...from, row: null, rowAbsolute: false },
    to: { ...to, row: null, rowAbsolute: false },
  })),
  fc.tuple(cornerArb, cornerArb).map(([from, to]) => ({
    pair: true as const,
    from: { ...from, column: null, columnAbsolute: false },
    to: { ...to, column: null, columnAbsolute: false },
  })),
);

// Every family of sheet name findReferences/sheetPrefix have to agree on:
// plain, needing quotes for a space/symbol, an apostrophe to double, a
// leading digit or dot, address-like, R1C1-shaped, and Unicode outside the
// BMP (the astral regression above, folded into the general sweep too).
const sheetNameArb = fc.oneof(
  fc.constant(""),
  fc.constantFrom(
    "Model",
    "Data_2025",
    "P&L 2025",
    "Bob's",
    "O'Brien, Ltd",
    "Q1 (EU)",
    "2025",
    ".hidden",
    "Q1",
    "XFD1048576",
    "R1C1",
    "Sheet!1",
    "Pārskats",
    "销售数据",
    "Q1-2025",
    "\u{1D400}",
    "Model\u{1D400}2",
    "\u{1D400}\u{1D401}\u{1D402}",
  ),
);

// [Book1.xlsx]Sheet1! - readPrefix's own workbook+name concatenation.
const workbookNameArb = fc.constantFrom(
  "[Book1.xlsx]Sheet1",
  "[Q1 model.xlsx]Data",
);

function tokenFor(
  sheet: string,
  shape: { pair: boolean; from: RefPart; to: RefPart },
): string {
  const prefix = sheetPrefix(sheet);
  const from = formatPart(shape.from);
  return shape.pair
    ? `${prefix}${from}:${formatPart(shape.to)}`
    : prefix + from;
}

describe("findReferences / formatReference round trip", () => {
  it("reads back exactly the token sheetPrefix + formatPart wrote, bare formula", () => {
    fc.assert(
      fc.property(sheetNameArb, refShapeArb, (sheet, shape) => {
        const token = tokenFor(sheet, shape);
        const formula = `=${token}`;
        const refs = findReferences(formula);
        expect(refs).toHaveLength(1);
        const [ref] = refs;
        expect(ref?.text).toBe(token);
        expect(ref?.sheet).toBe(sheet);
        expect(ref?.pair).toBe(shape.pair);
        expect(ref?.from).toEqual(shape.from);
        expect(ref?.to).toEqual(shape.to);
        expect(formatReference(ref!, ref!.from, ref!.to)).toBe(token);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("reads back the same token wrapped inside a function call and after an operator", () => {
    fc.assert(
      fc.property(sheetNameArb, refShapeArb, (sheet, shape) => {
        const token = tokenFor(sheet, shape);
        const wrapped = findReferences(`=SUM(${token})`);
        expect(wrapped).toHaveLength(1);
        expect(wrapped[0]?.text).toBe(token);

        const trailing = findReferences(`=${token}+1`);
        expect(trailing).toHaveLength(1);
        expect(trailing[0]?.text).toBe(token);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("round trips [Book]Sheet! external references the same way", () => {
    fc.assert(
      fc.property(workbookNameArb, refShapeArb, (sheet, shape) => {
        const token = tokenFor(sheet, shape);
        const [ref] = findReferences(`=${token}`);
        expect(ref?.text).toBe(token);
        expect(ref?.sheet).toBe(sheet);
        expect(formatReference(ref!, ref!.from, ref!.to)).toBe(token);
      }),
      { seed: SEED, numRuns: 500 },
    );
  });
});

describe("sheetPrefix doubles apostrophes and round trips through findReferences", () => {
  it("recovers the exact name it was given, for every generated sheet name", () => {
    fc.assert(
      fc.property(sheetNameArb, (name) => {
        if (name === "") {
          expect(sheetPrefix(name)).toBe("");
          return;
        }
        const [ref] = findReferences(`=${sheetPrefix(name)}A1`);
        expect(ref?.sheet).toBe(name);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("always doubles every apostrophe when it quotes", () => {
    fc.assert(
      fc.property(
        fc.stringMatching(/^[\p{L}\p{N} '&()._-]{1,20}$/u),
        (name) => {
          const prefix = sheetPrefix(name);
          if (!prefix.startsWith("'")) return; // written bare: nothing to double
          // Strip the delimiter quotes themselves (leading "'", trailing
          // "'!") before counting: counting raw "''" pairs across the WHOLE
          // prefix double-counts whenever the escaped content starts or ends
          // on an apostrophe, because that run of quote characters butts
          // straight up against the opening or closing delimiter.
          const inner = prefix.slice(1, -2);
          expect(inner.replaceAll("''", "'")).toBe(name);
          const realApostrophes = (name.match(/'/g) ?? []).length;
          expect((inner.match(/'/g) ?? []).length).toBe(2 * realApostrophes);
          // And it still reads back to the exact original name.
          const [ref] = findReferences(`=${prefix}A1`);
          expect(ref?.sheet).toBe(name);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
