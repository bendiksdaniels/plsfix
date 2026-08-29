// Prepare for sharing: the tidy pass a model gets before it leaves the desk.
// Every visible sheet goes back to A1 and the workbook lands on the first of
// them, and one scan reports what a reader would still find - hidden sheets,
// formulas pointing at other workbooks, names left on #REF! and autocolor still
// recoloring every edit. Nothing is deleted and no hidden sheet is touched.
// Zoom is not in the Office.js worksheet surface (pageLayout.zoom is print
// zoom), so the pane says so rather than pretending to reset it.

import { autocolorOnEditActive } from "./autocolor";
import {
  brokenIn,
  loadNames,
  pickScannableSheets,
  type ScannedSheet,
  SELECTION_CELL_CAP,
} from "./internal";
import { cellAddress } from "../find";
import { type CellValue } from "../model";
import {
  isExternalFormula,
  type ShareIssue,
  type ShareLink,
  shareReport,
} from "../share";

export interface ShareOptions {
  // The used range a sheet may have before it is skipped instead of read.
  // Defaults to the selection cap; the tests use a smaller one.
  maxCells?: number;
}

export interface ShareResult {
  report: ShareIssue[];
  // Visible sheets put back at A1, which is what the summary line counts.
  touchedSheets: number;
}

// The grid is read from the used range, so a hit's coordinates are relative to
// its top-left corner; the address the reader is given is the absolute one.
function externalLinks(sheets: ScannedSheet[]): ShareLink[] {
  const links: ShareLink[] = [];
  for (const sheet of sheets) {
    const formulas = sheet.range.formulas as CellValue[][];
    formulas.forEach((row, rowOffset) => {
      row.forEach((formula, colOffset) => {
        if (!isExternalFormula(formula)) return;
        links.push({
          sheet: sheet.name,
          address: cellAddress(
            sheet.range.rowIndex + rowOffset,
            sheet.range.columnIndex + colOffset,
          ),
          formula,
        });
      });
    });
  }
  return links;
}

// Excel selects only on the sheet it is showing, so each sheet is activated
// before its A1 is selected. Hidden sheets are left exactly as they are, and
// the run ends on the first visible sheet: the one the reader should open on.
function resetToA1(sheets: Excel.Worksheet[]): number {
  const visible = sheets.filter(
    (sheet) => sheet.visibility === Excel.SheetVisibility.visible,
  );
  for (const sheet of visible) {
    sheet.activate();
    sheet.getRange("A1").select();
  }
  visible[0]?.activate();
  return visible.length;
}

// Four syncs, one per phase, never one per sheet: the sheet list and the names,
// then every used range's extent, then the formulas of the sheets small enough
// to read, then the whole A1 reset in a single batch.
export async function prepareForSharing(
  options: ShareOptions = {},
): Promise<ShareResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    const names = loadNames(context);
    await context.sync();

    const used = sheets.items.map((sheet) =>
      sheet.getUsedRangeOrNullObject(true),
    );
    for (const range of used) {
      range.load("isNullObject,cellCount,rowIndex,columnIndex");
    }
    await context.sync();

    const { scanned, skippedSheets } = pickScannableSheets(
      sheets.items,
      used,
      options.maxCells ?? SELECTION_CELL_CAP,
    );
    for (const sheet of scanned) sheet.range.load("formulas");
    await context.sync();

    const report = shareReport({
      sheets: sheets.items.map((sheet) => ({
        name: sheet.name,
        visibility: sheet.visibility,
      })),
      externalLinks: externalLinks(scanned),
      brokenNames: brokenIn(names),
      skippedSheets,
      autocolorOnEdit: autocolorOnEditActive(),
    });

    const touchedSheets = resetToA1(sheets.items);
    await context.sync();
    return { report, touchedSheets };
  });
}
