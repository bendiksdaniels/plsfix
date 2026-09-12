// Clean past the data: the ballast a saved workbook carries in its used range.
// Owns the comparison of the sheet's used range against the one its values
// fill, and what is done with the difference - deleted, or only stripped of
// formatting when a chart or a shape sits on the sheet and a delete would move
// it. Runs outside pls,fix Undo, like the size cycles in sizes.ts: whole rows
// and columns are sheet state and getCellProperties carries none of it.

import { hostSupports } from "./internal";
import { syncWrite } from "./protection";

const STAGE = "Clean past the data";
// worksheet.charts.getCount is ExcelApi 1.4, worksheet.shapes.getCount 1.9.
const CHART_COUNT_API_SET = "1.4";
const SHAPE_COUNT_API_SET = "1.9";

interface Extent {
  lastRow: number;
  lastColumn: number;
}

function extentOf(range: Excel.Range): Extent {
  return {
    lastRow: range.rowIndex + range.rowCount - 1,
    lastColumn: range.columnIndex + range.columnCount - 1,
  };
}

function plural(count: number, noun: string): string {
  return `${count.toLocaleString()} ${noun}${count === 1 ? "" : "s"}`;
}

/** "1,204 rows and 12 columns", with whichever half there is. */
function removedPhrase(rows: number, columns: number): string {
  const parts: string[] = [];
  if (rows > 0) parts.push(plural(rows, "row"));
  if (columns > 0) parts.push(plural(columns, "column"));
  return parts.join(" and ");
}

/**
 * The tidy pass. Answers the line the toast shows: nothing to clean, nothing
 * past the data, the formats cleared, or what was removed.
 */
export async function cleanPastData(): Promise<string> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(false);
    const filled = sheet.getUsedRangeOrNullObject(true);
    for (const range of [used, filled]) {
      range.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    }
    // ClientResults: their value arrives with this same sync.
    const charts = hostSupports(CHART_COUNT_API_SET)
      ? sheet.charts.getCount()
      : null;
    const shapes = hostSupports(SHAPE_COUNT_API_SET)
      ? sheet.shapes.getCount()
      : null;
    await context.sync();

    // Nothing at all, or nothing but formatting: there is no data for the
    // surplus to be past, and a sheet dressed for numbers that have not been
    // typed yet is not ballast.
    if (used.isNullObject || filled.isNullObject) {
      return "Nothing on this sheet to clean.";
    }

    const edge = extentOf(used);
    const data = extentOf(filled);
    const rows = edge.lastRow - data.lastRow;
    const columns = edge.lastColumn - data.lastColumn;
    if (rows <= 0 && columns <= 0) {
      return `Nothing past the data on ${sheet.name}.`;
    }

    const drawings = (charts?.value ?? 0) + (shapes?.value ?? 0);
    const surplus = surplusRanges(sheet, used, data, rows, columns);
    if (drawings > 0) {
      for (const range of surplus) range.clear(Excel.ClearApplyTo.formats);
      await syncWrite(context, STAGE);
      return `Cleared the formats past the data on ${sheet.name}; rows and columns kept because the sheet has charts or shapes`;
    }

    // Rows first, then columns: the surplus columns sit beside the data, so
    // deleting the rows below it never moves them.
    if (rows > 0) {
      sheet
        .getRangeByIndexes(data.lastRow + 1, 0, rows, 1)
        .getEntireRow()
        .delete(Excel.DeleteShiftDirection.up);
    }
    if (columns > 0) {
      sheet
        .getRangeByIndexes(0, data.lastColumn + 1, 1, columns)
        .getEntireColumn()
        .delete(Excel.DeleteShiftDirection.left);
    }
    await syncWrite(context, STAGE);
    return `Removed ${removedPhrase(rows, columns)} past the data on ${sheet.name}`;
  });
}

// The blocks the surplus occupies inside the used range: the band under the
// data, and the band beside it. Bounded by the used range rather than by the
// sheet, so a format clear never reaches a million untouched cells.
function surplusRanges(
  sheet: Excel.Worksheet,
  used: Excel.Range,
  data: Extent,
  rows: number,
  columns: number,
): Excel.Range[] {
  const ranges: Excel.Range[] = [];
  if (rows > 0) {
    // The whole width of the used range: the band under the data covers the
    // surplus columns too, so the band beside it only needs the rows above.
    ranges.push(
      sheet.getRangeByIndexes(
        data.lastRow + 1,
        used.columnIndex,
        rows,
        used.columnCount,
      ),
    );
  }
  if (columns > 0) {
    const height = data.lastRow - used.rowIndex + 1;
    ranges.push(
      sheet.getRangeByIndexes(
        used.rowIndex,
        data.lastColumn + 1,
        height,
        columns,
      ),
    );
  }
  return ranges;
}
