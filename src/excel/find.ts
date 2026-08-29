// Super Find: one pass over the whole workbook - every sheet's used range, the
// defined names and the sheet names themselves - and the jump that follows a
// result back to its cell. Three syncs, one per phase, never one per sheet: the
// sheet list and names, then every used range's extent, then the grids of the
// sheets small enough to read.

import {
  cellAddress,
  includesQuery,
  matchCells,
  type MatchOptions,
  type RankedHit,
  NAME_ORDER,
  rankHits,
  SHEET_NAME_ROW,
} from "../find";
import {
  pickScannableSheets,
  type ScannedSheet,
  SELECTION_CELL_CAP,
} from "./internal";
import { ANCHOR_PREFIX } from "../link/model";
import { type CellValue } from "../model";

export interface FindHit {
  kind: "cell" | "name" | "sheet";
  sheet: string;
  address: string;
  text: string;
}

export interface FindOptions extends MatchOptions {
  // The used range a sheet may have before it is skipped instead of read.
  // Defaults to the selection cap; the tests use a smaller one.
  maxCells?: number;
}

export interface FindResult {
  hits: FindHit[];
  // Sheets whose used range was too large to read, named so the pane can say
  // the answer is incomplete rather than quietly leaving them out.
  skippedSheets: string[];
}

interface ScanRow extends RankedHit {
  hit: FindHit;
}

// A link anchor is ours, not the modeller's: the Links tab owns those names and
// a search of the model should never surface them.
function nameHits(
  items: Excel.NamedItem[],
  query: string,
  options: FindOptions,
): ScanRow[] {
  const rows: ScanRow[] = [];
  items.forEach((item, order) => {
    if (item.name.startsWith(ANCHOR_PREFIX)) return;
    const formula = typeof item.formula === "string" ? item.formula : "";
    const matched =
      includesQuery(item.name, query, options.matchCase) ||
      includesQuery(formula, query, options.matchCase);
    if (!matched) return;

    rows.push({
      sheetIndex: NAME_ORDER,
      row: order,
      col: 0,
      hit: { kind: "name", sheet: "", address: item.name, text: formula },
    });
  });
  return rows;
}

function sheetHits(
  items: Excel.Worksheet[],
  query: string,
  options: FindOptions,
): ScanRow[] {
  const rows: ScanRow[] = [];
  items.forEach((sheet, index) => {
    if (!includesQuery(sheet.name, query, options.matchCase)) return;
    rows.push({
      sheetIndex: index,
      row: SHEET_NAME_ROW,
      col: SHEET_NAME_ROW,
      hit: {
        kind: "sheet",
        sheet: sheet.name,
        address: "A1",
        text: sheet.name,
      },
    });
  });
  return rows;
}

// The grid is read from the used range, so a hit's coordinates are relative to
// its top-left corner; the address the modeller jumps to is the absolute one.
function cellHits(
  sheets: ScannedSheet[],
  query: string,
  options: FindOptions,
): ScanRow[] {
  const rows: ScanRow[] = [];
  for (const sheet of sheets) {
    const grid = {
      values: sheet.range.values as CellValue[][],
      formulas: sheet.range.formulas as CellValue[][],
    };
    for (const hit of matchCells(grid, query, options)) {
      const row = sheet.range.rowIndex + hit.row;
      const col = sheet.range.columnIndex + hit.col;
      rows.push({
        sheetIndex: sheet.index,
        row,
        col,
        hit: {
          kind: "cell",
          sheet: sheet.name,
          address: cellAddress(row, col),
          text: hit.text,
        },
      });
    }
  }
  return rows;
}

// Hidden sheets are searched as well: a number that moved is usually hiding on
// one, and the jump is what tells the modeller the sheet is out of reach.
export async function findInWorkbook(
  query: string,
  options: FindOptions,
): Promise<FindResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const names = context.workbook.names;
    names.load("items/name,items/formula");
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
    for (const sheet of scanned) sheet.range.load("values,formulas");
    await context.sync();

    const rows = [
      ...nameHits(names.items, query, options),
      ...sheetHits(sheets.items, query, options),
      ...cellHits(scanned, query, options),
    ];
    return { hits: rankHits(rows).map((row) => row.hit), skippedSheets };
  });
}

// Excel cannot select on a sheet it is not showing, so the jump activates first
// and refuses a hidden sheet by name rather than letting the host throw.
export async function jumpToHit(hit: FindHit): Promise<void> {
  await Excel.run(async (context) => {
    const range =
      hit.kind === "name"
        ? context.workbook.names.getItem(hit.address).getRange()
        : context.workbook.worksheets.getItem(hit.sheet).getRange(hit.address);
    const sheet = range.worksheet;
    sheet.load("name,visibility");
    await context.sync();

    if (sheet.visibility !== Excel.SheetVisibility.visible) {
      throw new Error(`${sheet.name} is hidden, so there is nowhere to jump.`);
    }
    sheet.activate();
    range.select();
    await context.sync();
  });
}
