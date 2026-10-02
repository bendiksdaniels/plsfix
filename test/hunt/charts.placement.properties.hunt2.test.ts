// Pass-2 property: placeChartBeside (src/excel/chart-place.ts) never hands
// back a box overlapping the source data or another chart, or off the
// sheet's grid - random anchors and existing charts, then anchors pinned to
// the sheet's last rows and columns.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  formatA1,
  installFakeHost,
} from "../fakehost";
import { placeChartBeside } from "../../src/excel/chart-place";
import { overlaps, type Box } from "../../src/layout";

enableStrictLoadSemantics();

const SEED = 20260927;
const RUNS = 200;

// chart-place.ts's own DEFAULT_ROW_HEIGHT/DEFAULT_COLUMN_WIDTH fallback,
// which is also what the fake reports for a cell nobody resized (verified
// against test/fakehost.ts's own DEFAULT_ROW_HEIGHT/DEFAULT_COLUMN_WIDTH):
// every box in this suite lives on that one shared grid, so a box built from
// row/column indices here lands exactly where the code under test computes it.
const ROW_HEIGHT = 15;
const COLUMN_WIDTH = 64;
const SHEET_ROWS = 1_048_576;
const SHEET_COLUMNS = 16_384;

let host: { helpers: FakeHelpers; workbook: FakeWorkbook };
let sheetCounter = 0;

function freshSheet(): string {
  const name = `S${String(sheetCounter)}`;
  sheetCounter += 1;
  host.helpers.addSheet(name);
  return name;
}

// Writes a value into every cell of the anchor rectangle, so the sheet's own
// used range is exactly the anchor - the same shape a real selection full of
// drivers or valuation methods leaves for chart-place.ts to read back.
function seedAnchor(
  sheetName: string,
  anchor: { row: number; col: number; rowCount: number; colCount: number },
): void {
  const grid = Array.from({ length: anchor.rowCount }, () =>
    Array.from({ length: anchor.colCount }, () => 1),
  );
  host.helpers.seed(`${sheetName}!${formatA1(anchor)}`, grid);
}

function anchorBox(anchor: {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
}): Box {
  return {
    left: anchor.col * COLUMN_WIDTH,
    top: anchor.row * ROW_HEIGHT,
    width: anchor.colCount * COLUMN_WIDTH,
    height: anchor.rowCount * ROW_HEIGHT,
  };
}

describe("placeChartBeside never overlaps the data or another chart", () => {
  const anchorArb = fc.record({
    row: fc.integer({ min: 0, max: 300 }),
    col: fc.integer({ min: 0, max: 80 }),
    rowCount: fc.integer({ min: 1, max: 5 }),
    colCount: fc.integer({ min: 1, max: 5 }),
  });
  const chartSizeArb = fc.record({
    width: fc.integer({ min: 100, max: 700 }),
    height: fc.integer({ min: 80, max: 450 }),
  });
  // Offsets from the anchor's own top-left corner, wide enough to land an
  // existing chart squarely on the right-of, below-anchor or below-everything
  // candidate at least some of the time, and just as often nowhere near them.
  const existingChartArb = fc.record({
    dLeft: fc.integer({ min: -300, max: 900 }),
    dTop: fc.integer({ min: -200, max: 1200 }),
    width: fc.integer({ min: 50, max: 400 }),
    height: fc.integer({ min: 40, max: 300 }),
  });
  const caseArb = fc.record({
    anchor: anchorArb,
    chartSize: chartSizeArb,
    existing: fc.array(existingChartArb, { minLength: 0, maxLength: 3 }),
  });

  it("holds over random anchors, sizes and pre-existing charts", async () => {
    host = installFakeHost({ sheets: [] });
    sheetCounter = 0;

    await fc.assert(
      fc.asyncProperty(caseArb, async ({ anchor, chartSize, existing }) => {
        const sheetName = freshSheet();
        seedAnchor(sheetName, anchor);
        const source = anchorBox(anchor);

        const existingBoxes: Box[] = existing.map((offset) => ({
          left: Math.max(0, source.left + offset.dLeft),
          top: Math.max(0, source.top + offset.dTop),
          width: offset.width,
          height: offset.height,
        }));
        for (const box of existingBoxes) {
          host.helpers.addChart(sheetName, { ...box });
        }

        // Captured after the pre-existing charts land, so the next (and
        // only) chart Excel.run adds below is at exactly this index.
        const before = host.workbook.charts.length;
        const placed = await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(sheetName);
          const range = sheet.getRangeByIndexes(
            anchor.row,
            anchor.col,
            anchor.rowCount,
            anchor.colCount,
          );
          const chart = sheet.charts.add(
            Excel.ChartType.columnClustered,
            range,
            Excel.ChartSeriesBy.auto,
          );
          chart.width = chartSize.width;
          chart.height = chartSize.height;
          await context.sync();
          return placeChartBeside(context, sheet, chart, range);
        });

        const newChart = host.workbook.charts[before];
        if (!newChart) throw new Error("hunt2: the new chart never landed");
        if (!placed) return;

        const result: Box = {
          left: newChart.left ?? 0,
          top: newChart.top ?? 0,
          width: chartSize.width,
          height: chartSize.height,
        };
        expect(overlaps(result, source)).toBe(false);
        for (const box of existingBoxes) {
          expect(overlaps(result, box)).toBe(false);
        }
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("placeChartBeside when the sheet's used range reaches past the anchor", () => {
  // Every earlier case seeds the used range as exactly the anchor; here one
  // more cell lands well below it, at the anchor's own column, so
  // "below everything" (the sheet's used range) is a genuinely different
  // spot from "below the anchor" - the two candidates chart-place.ts's own
  // header calls out by name - and a chart parked right under the anchor is
  // what forces the walk from one to the other instead of skipping it.
  const caseArb = fc.record({
    row: fc.integer({ min: 0, max: 100 }),
    col: fc.integer({ min: 0, max: 60 }),
    rowCount: fc.integer({ min: 1, max: 4 }),
    colCount: fc.integer({ min: 1, max: 4 }),
    extraRowsBelow: fc.integer({ min: 1, max: 12 }),
    width: fc.integer({ min: 100, max: 600 }),
    height: fc.integer({ min: 60, max: 400 }),
    blockBelowAnchor: fc.boolean(),
  });

  it("never lands on the data below the anchor, or on a chart parked right under it", async () => {
    host = installFakeHost({ sheets: [] });
    sheetCounter = 0;

    await fc.assert(
      fc.asyncProperty(
        caseArb,
        async ({
          row,
          col,
          rowCount,
          colCount,
          extraRowsBelow,
          width,
          height,
          blockBelowAnchor,
        }) => {
          const sheetName = freshSheet();
          const anchor = { row, col, rowCount, colCount };
          seedAnchor(sheetName, anchor);
          const source = anchorBox(anchor);

          // One more cell, well clear of the anchor's own rows, pushes the
          // sheet's used range past it - candidateCorners' own "below
          // everything" reads this back as sheetBottom.
          const extraRow = row + rowCount + extraRowsBelow - 1;
          host.helpers.seed(
            `${sheetName}!${formatA1({ row: extraRow, col, rowCount: 1, colCount: 1 })}`,
            [[1]],
          );
          const usedBottomRow = extraRow + 1;
          const usedBox: Box = {
            left: col * COLUMN_WIDTH,
            top: row * ROW_HEIGHT,
            width: colCount * COLUMN_WIDTH,
            height: (usedBottomRow - row) * ROW_HEIGHT,
          };

          const columns = Math.max(1, Math.ceil(width / COLUMN_WIDTH));
          const belowColumn = Math.max(
            0,
            Math.min(col, SHEET_COLUMNS - columns),
          );
          const existingBoxes: Box[] = [];
          if (blockBelowAnchor) {
            // Sitting exactly on the "below the anchor" candidate: pick()
            // must step past it to "below everything" instead of stopping.
            existingBoxes.push({
              left: belowColumn * COLUMN_WIDTH,
              top: (row + rowCount + 1) * ROW_HEIGHT,
              width,
              height,
            });
          }
          for (const box of existingBoxes) {
            host.helpers.addChart(sheetName, { ...box });
          }

          const before = host.workbook.charts.length;
          const placed = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getRangeByIndexes(row, col, rowCount, colCount);
            const chart = sheet.charts.add(
              Excel.ChartType.columnClustered,
              range,
              Excel.ChartSeriesBy.auto,
            );
            chart.width = width;
            chart.height = height;
            await context.sync();
            return placeChartBeside(context, sheet, chart, range);
          });

          const newChart = host.workbook.charts[before];
          if (!newChart) throw new Error("hunt2: the new chart never landed");
          if (!placed) return;

          const result: Box = {
            left: newChart.left ?? 0,
            top: newChart.top ?? 0,
            width,
            height,
          };
          expect(overlaps(result, source)).toBe(false);
          expect(overlaps(result, usedBox)).toBe(false);
          for (const box of existingBoxes) {
            expect(overlaps(result, box)).toBe(false);
          }
        },
      ),
      { seed: SEED, numRuns: 150 },
    );
  });
});

describe("placeChartBeside at the sheet's last rows and columns", () => {
  const edgeArb = fc.constantFrom("row", "column", "corner");
  const caseArb = fc.record({
    edge: edgeArb,
    offset: fc.integer({ min: 0, max: 8 }),
    rowCount: fc.integer({ min: 1, max: 4 }),
    colCount: fc.integer({ min: 1, max: 4 }),
    width: fc.integer({ min: 64, max: 700 }),
    height: fc.integer({ min: 15, max: 450 }),
  });

  it("never asks for a box the grid cannot hold, or throws", async () => {
    host = installFakeHost({ sheets: [] });
    sheetCounter = 0;

    await fc.assert(
      fc.asyncProperty(
        caseArb,
        async ({ edge, offset, rowCount, colCount, width, height }) => {
          const sheetName = freshSheet();
          const pinnedRow = Math.max(0, SHEET_ROWS - rowCount - offset);
          const pinnedCol = Math.max(0, SHEET_COLUMNS - colCount - offset);
          const anchor = {
            row: edge === "column" ? 10 : pinnedRow,
            col: edge === "row" ? 10 : pinnedCol,
            rowCount,
            colCount,
          };
          seedAnchor(sheetName, anchor);
          const source = anchorBox(anchor);

          const before = host.workbook.charts.length;
          const placed = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getRangeByIndexes(
              anchor.row,
              anchor.col,
              anchor.rowCount,
              anchor.colCount,
            );
            const chart = sheet.charts.add(
              Excel.ChartType.columnClustered,
              range,
              Excel.ChartSeriesBy.auto,
            );
            chart.width = width;
            chart.height = height;
            await context.sync();
            return placeChartBeside(context, sheet, chart, range);
          });

          const newChart = host.workbook.charts[before];
          if (!newChart) throw new Error("hunt2: the new chart never landed");
          if (!placed) return;

          const result: Box = {
            left: newChart.left ?? 0,
            top: newChart.top ?? 0,
            width,
            height,
          };
          expect(overlaps(result, source)).toBe(false);
          // The boxes this suite builds only ever move down or right of a
          // grid-aligned corner, so a box that fits at all fits inside the
          // grid's own last row and column, in points.
          expect(result.top + result.height).toBeLessThanOrEqual(
            SHEET_ROWS * ROW_HEIGHT,
          );
          expect(result.left + result.width).toBeLessThanOrEqual(
            SHEET_COLUMNS * COLUMN_WIDTH,
          );
        },
      ),
      { seed: SEED, numRuns: 120 },
    );
  });
});
