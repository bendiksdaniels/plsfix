// The unused-style scrubber: every custom cell style in the workbook, matched
// against the styles the cells themselves are wearing. One sync per phase - the
// style table and the sheet list, then every used range's extent, then the
// grids of the sheets small enough to read - never one sync per sheet.

import {
  pickScannableSheets,
  SCAN_CELL_CAP,
  SELECTION_CELL_CAP,
} from "./internal";
import { unusedStyles, type WorkbookStyle } from "../styles-audit";

export interface StyleScan {
  // The custom styles no cell wears, sorted.
  unused: string[];
  // Every style in the workbook, built-ins included: what the count is out of.
  total: number;
  // Sheets whose used range was too large to read. A style worn only there
  // would look unused, so the pane says the answer is incomplete and the
  // delete refuses to run.
  skippedSheets: string[];
}

type CellGrid = OfficeExtension.ClientResult<Excel.CellProperties[][]>;

// Which sheets are worth reading is the same question Super Find and Prepare
// for sharing ask, so it is asked in the same place: an empty sheet dresses no
// cells, and a sheet - or a run of them - past the cap would overflow the
// request payload, so it is reported as skipped rather than half-read. Style
// properties are the heaviest read of the three, which is why the cap matters
// most here.
function readStyles(
  sheets: Excel.Worksheet[],
  ranges: Excel.Range[],
  cap: number,
  totalCap: number,
): { grids: CellGrid[]; skippedSheets: string[] } {
  const { scanned, skippedSheets } = pickScannableSheets(
    sheets,
    ranges,
    cap,
    totalCap,
  );
  return {
    grids: scanned.map((sheet) =>
      sheet.range.getCellProperties({ style: true }),
    ),
    skippedSheets,
  };
}

function wornStyles(grids: CellGrid[]): Set<string> {
  const worn = new Set<string>();
  for (const grid of grids) {
    for (const row of grid.value) {
      for (const cell of row) {
        if (typeof cell.style === "string") worn.add(cell.style);
      }
    }
  }
  return worn;
}

// Hidden sheets are read as well: a style worn only on a back-room sheet is in
// use, and deleting it would restyle those cells behind the modeller's back.
async function scanStyles(
  context: Excel.RequestContext,
  cap: number,
  totalCap: number,
): Promise<StyleScan> {
  const styles = context.workbook.styles;
  styles.load("items/name,items/builtIn");
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const all: WorkbookStyle[] = styles.items.map((style) => ({
    name: style.name,
    builtIn: style.builtIn,
  }));
  const ranges = sheets.items.map((sheet) =>
    sheet.getUsedRangeOrNullObject(true),
  );
  for (const range of ranges) range.load("isNullObject,cellCount");
  await context.sync();

  const { grids, skippedSheets } = readStyles(
    sheets.items,
    ranges,
    cap,
    totalCap,
  );
  await context.sync();

  return {
    unused: unusedStyles(all, wornStyles(grids)),
    total: all.length,
    skippedSheets,
  };
}

export async function listUnusedStyles(
  maxCells = SELECTION_CELL_CAP,
  maxTotalCells = SCAN_CELL_CAP,
): Promise<StyleScan> {
  return Excel.run((context) => scanStyles(context, maxCells, maxTotalCells));
}

// Irreversible, and every cell wearing a deleted style is restyled with it, so
// the list is re-derived here rather than trusted from the pane; a workbook
// with a sheet too large to read is refused outright, because a style could
// well be in use on it.
export async function deleteUnusedStyles(
  names: string[],
  maxCells = SELECTION_CELL_CAP,
  maxTotalCells = SCAN_CELL_CAP,
): Promise<number> {
  return Excel.run(async (context) => {
    const scan = await scanStyles(context, maxCells, maxTotalCells);
    if (scan.skippedSheets.length > 0) {
      throw new Error("styles: some sheets were too large to scan");
    }

    const unused = new Set(scan.unused);
    const doomed = names.filter((name) => unused.has(name));
    const styles = context.workbook.styles;
    for (const name of doomed) styles.getItem(name).delete();
    await context.sync();
    return doomed.length;
  });
}
