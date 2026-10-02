// Pass-2 property: duplicateFormula's two algebraic guarantees, independent of
// the inside/outside rule's own internals - a zero-offset same-sheet
// "duplicate" changes nothing, and moving a block's formula by an offset then
// by its negative returns exactly to where it started. Hammered with
// fast-check over random blocks, offsets and sheet-qualified references, both
// inside and outside the copied block.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  duplicateFormula,
  type CellBlock,
  type CellOffset,
} from "../../src/formula-duplicate";
import { formatPart, sheetPrefix, type RefPart } from "../../src/formula-refs";

const SEED = 20260927;
const RUNS = 1500;

// Kept well inside the grid (GRID_ROWS ~1.05M, GRID_COLUMNS 16,384) with
// enough headroom either side of the largest offset, so no reference in these
// tests ever falls off the edge - that "stays put" rule is formula-duplicate
// .test.ts's own territory, not this property's.
const blockArb: fc.Arbitrary<CellBlock> = fc.record({
  sheet: fc.constantFrom("Model", "Data", "O'Brien", "P&L 2025"),
  row: fc.integer({ min: 200, max: 800 }),
  column: fc.integer({ min: 20, max: 80 }),
  rowCount: fc.integer({ min: 1, max: 6 }),
  columnCount: fc.integer({ min: 1, max: 6 }),
});

const offsetArb: fc.Arbitrary<CellOffset> = fc.record({
  rows: fc.integer({ min: -80, max: 80 }),
  columns: fc.integer({ min: -15, max: 15 }),
});

function part(
  row: number,
  column: number,
  rowAbs: boolean,
  colAbs: boolean,
): RefPart {
  return { row, column, rowAbsolute: rowAbs, columnAbsolute: colAbs };
}

// A formula made of several references: some inside the block (random cell
// within it), some outside on the block's own sheet (unqualified and
// qualified with the block's sheet name), and one on a third, unrelated
// sheet. absFlags controls which corners carry a $ - unrelated to the rule
// under test, just noise the rewriter has to carry through untouched.
// block.row/.column stay in [200,800]/[20,80], rowCount/columnCount <= 6 and
// the offsets this file uses stay within +-80 rows / +-15 columns, so the
// original block and any block it is moved to never reach past row 886 or
// column 101. 2000/2100 and columns 500/501 are fixed, well clear of that
// whole range, so an "outside" reference here can never coincide with either
// block by coincidence of the values fast-check happens to pick.
const OUTSIDE_ROW_A = 2000;
const OUTSIDE_ROW_B = 2100;
const OUTSIDE_COLUMN_A = 500;
const OUTSIDE_COLUMN_B = 501;

function buildFormula(
  block: CellBlock,
  insideOffset: { dr: number; dc: number },
  absA: boolean,
  absB: boolean,
): { formula: string; insideToken: string; elsewhereToken: string } {
  const insideRef = part(
    block.row + insideOffset.dr,
    block.column + insideOffset.dc,
    absA,
    absB,
  );
  const insideToken = formatPart(insideRef);
  const outsideUnqualified = formatPart(
    part(OUTSIDE_ROW_A, OUTSIDE_COLUMN_A, false, false),
  );
  const outsideQualified =
    sheetPrefix(block.sheet) +
    formatPart(part(OUTSIDE_ROW_B, OUTSIDE_COLUMN_B, false, false));
  const elsewhereToken =
    sheetPrefix("Elsewhere") + formatPart(part(9, 25, true, true));
  const formula = `=${insideToken}+${outsideUnqualified}+${outsideQualified}+${elsewhereToken}`;
  return { formula, insideToken, elsewhereToken };
}

const insideOffsetArb = (block: CellBlock) =>
  fc.record({
    dr: fc.integer({ min: 0, max: block.rowCount - 1 }),
    dc: fc.integer({ min: 0, max: block.columnCount - 1 }),
  });

describe("duplicateFormula: zero offset, same sheet, is the identity", () => {
  it("changes nothing for a formula mixing inside, outside and elsewhere refs", () => {
    fc.assert(
      fc.property(
        blockArb.chain((block) =>
          fc.tuple(
            fc.constant(block),
            insideOffsetArb(block),
            fc.boolean(),
            fc.boolean(),
          ),
        ),
        ([block, insideOffset, absA, absB]) => {
          const { formula } = buildFormula(block, insideOffset, absA, absB);
          const zero: CellOffset = { rows: 0, columns: 0 };
          expect(duplicateFormula(formula, block, zero)).toBe(formula);
          expect(duplicateFormula(formula, block, zero, block.sheet)).toBe(
            formula,
          );
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("duplicateFormula: an offset and then its negative returns to the start", () => {
  it("round-trips an in-block reference through a same-sheet move and back", () => {
    fc.assert(
      fc.property(
        blockArb.chain((block) =>
          fc.tuple(
            fc.constant(block),
            insideOffsetArb(block),
            offsetArb,
            fc.boolean(),
            fc.boolean(),
          ),
        ),
        ([block, insideOffset, offset, absA, absB]) => {
          const { formula } = buildFormula(block, insideOffset, absA, absB);

          const moved = duplicateFormula(formula, block, offset);
          const movedBlock: CellBlock = {
            ...block,
            row: block.row + offset.rows,
            column: block.column + offset.columns,
          };
          const back = duplicateFormula(moved, movedBlock, {
            rows: -offset.rows,
            columns: -offset.columns,
          });
          expect(back).toBe(formula);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  // An IN-BLOCK reference is the only one this double cross can promise
  // anything about: it moves with the block both times, by construction, so
  // the two offsets and the two sheet crossings both cancel out exactly. An
  // OUTSIDE reference does not round-trip the same way - crossing sheets once
  // qualifies it (correctly: "Model!Z9" now names exactly where it points),
  // and that qualification is not undone by a second, unrelated crossing
  // back through a different block, because nothing records that the
  // qualifier was added rather than typed by hand. That is by design (the
  // module header: "a reference outside it keeps the cells it was written
  // for"), not a round trip this property should expect.
  it("round-trips a lone in-block reference through a cross-sheet move and back", () => {
    fc.assert(
      fc.property(
        blockArb.chain((block) =>
          fc.tuple(
            fc.constant(block),
            insideOffsetArb(block),
            offsetArb,
            fc.boolean(),
            fc.boolean(),
            fc.constantFrom("Data", "Model", "Q1", "O'Brien"),
          ),
        ),
        ([block, insideOffset, offset, absA, absB, toSheet]) => {
          fc.pre(toSheet.toLowerCase() !== block.sheet.toLowerCase());
          const insideRef = part(
            block.row + insideOffset.dr,
            block.column + insideOffset.dc,
            absA,
            absB,
          );
          const formula = `=${formatPart(insideRef)}`;

          const moved = duplicateFormula(formula, block, offset, toSheet);
          const movedBlock: CellBlock = {
            ...block,
            sheet: toSheet,
            row: block.row + offset.rows,
            column: block.column + offset.columns,
          };
          const back = duplicateFormula(
            moved,
            movedBlock,
            { rows: -offset.rows, columns: -offset.columns },
            block.sheet,
          );
          expect(back).toBe(formula);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("leaves an explicit cross-sheet reference alone once qualified, even through a later, unrelated block", () => {
    fc.assert(
      fc.property(
        blockArb.chain((block) =>
          fc.tuple(fc.constant(block), insideOffsetArb(block)),
        ),
        offsetArb,
        fc.boolean(),
        fc.boolean(),
        fc.constantFrom("Data", "Model", "Q1", "O'Brien"),
        ([block, insideOffset], offset, absA, absB, toSheet) => {
          fc.pre(toSheet.toLowerCase() !== block.sheet.toLowerCase());
          const { formula } = buildFormula(block, insideOffset, absA, absB);
          // outsideUnqualified crosses once here and becomes "<block.sheet>!...".
          const onceMoved = duplicateFormula(formula, block, offset, toSheet);
          const qualified =
            sheetPrefix(block.sheet) +
            formatPart(part(OUTSIDE_ROW_A, OUTSIDE_COLUMN_A, false, false));
          expect(onceMoved).toContain(qualified);

          // A second, unrelated block far from row 2000 must leave it alone,
          // whichever sheet that second paste lands on.
          const unrelatedBlock: CellBlock = {
            sheet: toSheet,
            row: 5,
            column: 5,
            rowCount: 2,
            columnCount: 2,
          };
          const again = duplicateFormula(
            onceMoved,
            unrelatedBlock,
            { rows: 1000, columns: 1 },
            "Q4",
          );
          expect(again).toContain(qualified);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("duplicateFormula: an elsewhere reference is never touched", () => {
  it("keeps a third-sheet reference byte for byte at any offset, same or cross sheet", () => {
    fc.assert(
      fc.property(
        blockArb.chain((block) =>
          fc.tuple(fc.constant(block), insideOffsetArb(block), offsetArb),
        ),
        fc.boolean(),
        fc.boolean(),
        fc.constantFrom("Data", "Model", "Q1"),
        ([block, insideOffset, offset], absA, absB, toSheet) => {
          const { formula, elsewhereToken } = buildFormula(
            block,
            insideOffset,
            absA,
            absB,
          );
          const sameSheetResult = duplicateFormula(formula, block, offset);
          const crossSheetResult = duplicateFormula(
            formula,
            block,
            offset,
            toSheet,
          );
          expect(sameSheetResult).toContain(elsewhereToken);
          expect(crossSheetResult).toContain(elsewhereToken);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
