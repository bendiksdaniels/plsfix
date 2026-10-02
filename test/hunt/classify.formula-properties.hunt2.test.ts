// Property: classifyCell's class ignores whitespace, function-name case and
// optional sheet quoting, and a cross-sheet or external reference outranks a
// hardcode in the same formula; auditGrid's "typed" mark matches an
// independently written rule over every small row shape.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { type AuditMark, auditGrid } from "../../src/audit";
import { type CellClass, classifyCell } from "../../src/classify";
import { type CellValue } from "../../src/model";

const SEED = 20260927;
const RUNS = 2000;

// ---------------------------------------------------------------------------
// References: same-sheet, a table, a defined name, cross-sheet and external,
// each with the class classifyCell must answer when nothing else is added.
// ---------------------------------------------------------------------------

interface RefCase {
  label: string;
  quoted: string;
  // null when Excel never permits an unquoted spelling (a name with a
  // hyphen, an ampersand, a space or a non-ASCII letter).
  unquoted: string | null;
  expected: CellClass;
}

function quoteName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

const PLAIN_SHEET_NAMES = ["Sheet1", "Data2", "Assumptions"];
const QUOTED_ONLY_SHEET_NAMES = [
  "Q1-2026", // hyphen
  "R&D", // ampersand
  "PiegÄdÄtÄji", // non-ASCII letters
  "My Sheet", // space
  "O'Brien's Model", // apostrophes, doubled when quoted
];
const PLAIN_BOOK_NAMES = ["Budget.xlsx", "Model.xlsx"];
// A closed workbook referenced by its index into Excel's own link table.
const BOOK_INDEXES = ["[1]", "[2]"];

const sameSheetCases: RefCase[] = [
  {
    label: "same-sheet cell",
    quoted: "A1",
    unquoted: "A1",
    expected: "formula",
  },
  {
    label: "same-sheet absolute range",
    quoted: "$B$2:$C$3",
    unquoted: "$B$2:$C$3",
    expected: "formula",
  },
  {
    label: "a table reference",
    quoted: "Table1[Revenue]",
    unquoted: "Table1[Revenue]",
    expected: "formula",
  },
  {
    label: "a defined name",
    quoted: "Revenue_2026",
    unquoted: "Revenue_2026",
    expected: "formula",
  },
];

const crossSheetCases: RefCase[] = [
  ...PLAIN_SHEET_NAMES.map((name): RefCase => ({
    label: `cross-sheet, plain name ${name}`,
    quoted: `${quoteName(name)}!A1`,
    unquoted: `${name}!A1`,
    expected: "crossSheet",
  })),
  ...QUOTED_ONLY_SHEET_NAMES.map((name): RefCase => ({
    label: `cross-sheet, quoted-only name ${name}`,
    quoted: `${quoteName(name)}!A1`,
    unquoted: null,
    expected: "crossSheet",
  })),
];

const externalCases: RefCase[] = [
  ...PLAIN_BOOK_NAMES.flatMap((book) =>
    PLAIN_SHEET_NAMES.map((name): RefCase => ({
      label: `external, plain [${book}]${name}`,
      quoted: `${quoteName(`[${book}]${name}`)}!A1`,
      unquoted: `[${book}]${name}!A1`,
      expected: "external",
    })),
  ),
  ...PLAIN_BOOK_NAMES.flatMap((book) =>
    QUOTED_ONLY_SHEET_NAMES.map((name): RefCase => ({
      label: `external, quoted-only [${book}]${name}`,
      quoted: `${quoteName(`[${book}]${name}`)}!A1`,
      unquoted: null,
      expected: "external",
    })),
  ),
  ...BOOK_INDEXES.flatMap((index) =>
    [...PLAIN_SHEET_NAMES, ...QUOTED_ONLY_SHEET_NAMES].map((name): RefCase => ({
      label: `external, closed workbook ${index}${name}`,
      quoted: `${quoteName(`${index}${name}`)}!A1`,
      // A closed workbook's index form is unquoted only when its sheet name
      // needs no quoting of its own.
      unquoted: PLAIN_SHEET_NAMES.includes(name) ? `${index}${name}!A1` : null,
      expected: "external",
    })),
  ),
];

const allRefCases = [...sameSheetCases, ...crossSheetCases, ...externalCases];
const refCaseArb = fc.constantFrom(...allRefCases);

// A formula's tail: nothing, an identity constant (0 and 1 never count as a
// hardcode, added or subtracted, front or back), or a real one.
type Tail = "none" | "plusOne" | "minusOneFront" | "timesZero" | "plusTwo";
const tailArb = fc.constantFrom<Tail>(
  "none",
  "plusOne",
  "minusOneFront",
  "timesZero",
  "plusTwo",
);

function withTail(ref: string, tail: Tail): string {
  switch (tail) {
    case "none":
      return ref;
    case "plusOne":
      return `${ref}+1`;
    case "minusOneFront":
      return `1-${ref}`;
    case "timesZero":
      return `${ref}*0`;
    case "plusTwo":
      return `${ref}+2`;
  }
}

// Only "plusTwo" carries a real hardcode; every other tail is either bare or
// an identity constant, which classifyCell's own rule (IDENTITY_CONSTANTS)
// strips before its digit test.
function expectedClass(base: CellClass, tail: Tail): CellClass {
  if (tail !== "plusTwo") return base;
  // external and crossSheet outrank a hardcode; only a same-sheet reference
  // (base "formula") can turn "partial".
  return base === "formula" ? "partial" : base;
}

// Wraps the reference in a function call, so whitespace around the comma and
// the function name's case are exercised the same way a modeller might type
// either. The identity constant sits in a second argument so the tail
// variants above apply to it unchanged (SUM's own arithmetic never runs -
// classifyCell reads the formula text, not its value).
function wrapped(
  ref: string,
  tail: Tail,
  functionCase: "SUM" | "sum" | "Sum",
  spaced: boolean,
): string {
  const inner = withTail(ref, tail);
  return spaced
    ? `=${functionCase}( ${inner} , 0 )`
    : `=${functionCase}(${inner},0)`;
}

const functionCaseArb = fc.constantFrom<"SUM" | "sum" | "Sum">(
  "SUM",
  "sum",
  "Sum",
);

describe("classifyCell over generated formulas", () => {
  it("never depends on a function name's case or extra whitespace around its arguments", () => {
    fc.assert(
      fc.property(
        refCaseArb,
        tailArb,
        functionCaseArb,
        functionCaseArb,
        (refCase, tail, caseA, caseB) => {
          const compact = wrapped(refCase.quoted, tail, caseA, false);
          const spaced = wrapped(refCase.quoted, tail, caseB, true);
          const expected = expectedClass(refCase.expected, tail);
          expect(classifyCell(compact, 1)).toBe(expected);
          expect(classifyCell(spaced, 1)).toBe(expected);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("never depends on quoting a sheet or workbook name Excel would also take bare", () => {
    fc.assert(
      fc.property(refCaseArb, tailArb, (refCase, tail) => {
        const expected = expectedClass(refCase.expected, tail);
        const quotedFormula = `=${withTail(refCase.quoted, tail)}`;
        expect(classifyCell(quotedFormula, 1)).toBe(expected);
        if (refCase.unquoted !== null) {
          const unquotedFormula = `=${withTail(refCase.unquoted, tail)}`;
          expect(classifyCell(unquotedFormula, 1)).toBe(expected);
        }
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("ranks a cross-sheet or external reference above the hardcode it also carries", () => {
    // Exhaustive over every generated reference, not a sample: "plusTwo" is
    // the one tail with a real, non-identity hardcode, and every one of
    // these must still answer its own class, never "partial".
    for (const refCase of [...crossSheetCases, ...externalCases]) {
      expect(classifyCell(`=${withTail(refCase.quoted, "plusTwo")}`, 3)).toBe(
        refCase.expected,
      );
    }
  });
});

// ---------------------------------------------------------------------------
// auditGrid's "typed" mark: a numeric cell reads as typed exactly when every
// EXISTING neighbour across is a formula and, where both exist, the SAME
// formula - independently re-derived here from the invariant's own words,
// not by re-reading src/audit.ts.
// ---------------------------------------------------------------------------

type Cell = "formulaA" | "formulaB" | "blank" | "text" | number;

const cellArb: fc.Arbitrary<Cell> = fc.oneof(
  fc.constant<Cell>("formulaA"),
  fc.constant<Cell>("formulaB"),
  fc.constant<Cell>("blank"),
  fc.constant<Cell>("text"),
  fc.integer({ min: -50, max: 50 }),
);

const rowArb = fc.array(cellArb, { minLength: 1, maxLength: 8 });

function isFormulaCell(cell: Cell | undefined): boolean {
  return cell === "formulaA" || cell === "formulaB";
}

function expectedTyped(row: Cell[], index: number): boolean {
  const left = index - 1 >= 0 ? row[index - 1] : undefined;
  const right = index + 1 < row.length ? row[index + 1] : undefined;
  if (left !== undefined && !isFormulaCell(left)) return false;
  if (right !== undefined && !isFormulaCell(right)) return false;
  if (left === undefined && right === undefined) return false;
  if (left !== undefined && right !== undefined) return left === right;
  return true;
}

function toGridCell(cell: Cell): CellValue {
  switch (cell) {
    case "formulaA":
      return "=AAA";
    case "formulaB":
      return "=BBB";
    case "blank":
      return null;
    case "text":
      return "label";
    default:
      return cell;
  }
}

describe("auditGrid's typed mark over generated rows", () => {
  it("matches an independently derived rule for every numeric cell's formula neighbours", () => {
    fc.assert(
      fc.property(rowArb, (row) => {
        const grid: CellValue[][] = [row.map(toGridCell)];
        const marks: AuditMark[] = auditGrid(grid)[0]!;
        row.forEach((cell, index) => {
          if (typeof cell !== "number") return;
          expect(marks[index] === "typed").toBe(expectedTyped(row, index));
        });
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
