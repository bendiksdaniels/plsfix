// Model check: one read-only pass over the workbook that lists what a reviewer
// would flag. It never paints and never touches an overlay's fill snapshot, and
// it activates nothing - a very hidden sheet is listed, never opened. The rules
// and the wording live in src/model-check.ts; this file only reads: the used
// range of every sheet in one load each, then the two scrubbers' own scans for
// the names and the style table.

import {
  brokenIn,
  loadNames,
  pickScannableSheets,
  type ScannedSheet,
} from "./internal";
import { listUnusedStyles } from "./styles";
import { type AuditMark, auditGrid } from "../audit";
import { cellAddress } from "../find";
import { type CellValue, isFormula } from "../model";
import {
  buildReport,
  type CheckKind,
  type Finding,
  hardcodeIn,
  type ModelCheckReport,
  volatileIn,
} from "../model-check";
import { isExternalFormula } from "../share";

// What one sheet's used range may hold before it is read at all. A sheet past
// it is named as skipped rather than half-read, and the same number bounds the
// whole pass: four grids of a sheet this size already make one of the largest
// requests the host will answer, and a workbook of them would overflow it.
export const MODEL_CHECK_CELL_CAP = 200_000;

interface CellRead {
  formula: CellValue;
  value: CellValue;
  type: Excel.RangeValueType;
  mark: AuditMark;
}

function finding(
  kind: CheckKind,
  sheet: string | null,
  ref: string | null,
  note: string,
  count = 1,
): Finding {
  return { kind, sheet, ref, count, note };
}

// One cell, every rule that has something to say about it. A cell can fail more
// than one: a hardcoded threshold inside a volatile lookup is two findings,
// because a reviewer would write down two lines.
function cellFindings(sheet: string, ref: string, cell: CellRead): Finding[] {
  const found: Finding[] = [];
  if (cell.type === Excel.RangeValueType.error) {
    found.push(finding("formulaError", sheet, ref, String(cell.value ?? "")));
  }
  if (!isFormula(cell.formula)) return found;

  const formula = cell.formula;
  if (isExternalFormula(formula)) {
    found.push(finding("externalLink", sheet, ref, formula));
  }
  if (hardcodeIn(formula)) {
    found.push(finding("hardcodeInFormula", sheet, ref, formula));
  }
  const volatile = volatileIn(formula);
  if (volatile !== null) {
    found.push(
      finding("volatileFormula", sheet, ref, `${volatile} in ${formula}`),
    );
  }
  // "lone" is the audit overlay's word for a formula whose neighbours are
  // formulas and none of them matches it: the deviation in a filled row.
  if (cell.mark === "lone") {
    found.push(finding("inconsistentFormula", sheet, ref, formula));
  }
  return found;
}

// The grids arrive relative to the used range's top-left corner, so a hit's
// address is built from the range's own row and column indexes.
function sheetFindings(sheet: ScannedSheet): Finding[] {
  const range = sheet.range;
  const formulas = range.formulas as CellValue[][];
  const values = range.values as CellValue[][];
  const types = range.valueTypes;
  const marks = auditGrid(range.formulasR1C1 as CellValue[][]);

  return formulas.flatMap((row, rowOffset) =>
    row.flatMap((formula, colOffset) =>
      cellFindings(
        sheet.name,
        cellAddress(range.rowIndex + rowOffset, range.columnIndex + colOffset),
        {
          formula,
          value: values[rowOffset]?.[colOffset] ?? null,
          type: types[rowOffset]?.[colOffset] ?? Excel.RangeValueType.empty,
          mark: marks[rowOffset]?.[colOffset] ?? "none",
        },
      ),
    ),
  );
}

// Very hidden is set outside Excel's UI and cannot be shown from one, so the
// reviewer is told which kind of hidden they are looking at. Neither kind is
// ever activated by this pass.
function hiddenSheetFindings(sheets: Excel.Worksheet[]): Finding[] {
  return sheets
    .filter((sheet) => sheet.visibility !== Excel.SheetVisibility.visible)
    .map((sheet) =>
      finding(
        "hiddenSheet",
        sheet.name,
        null,
        sheet.visibility === Excel.SheetVisibility.veryHidden
          ? "Very hidden, and can only be shown outside Excel"
          : "Hidden",
      ),
    );
}

// The style table is the scrubber's own scan, caps and all. A sheet it could
// not read could be wearing any of these styles, so the whole kind is dropped
// rather than reported half right: the Styles section says so on its own. A
// host that refuses the read outright gets the same answer - getCellProperties
// is the one call of this pass a host can say no to, and one style kind is not
// worth the other seven and the whole check with them.
async function unusedStyleFindings(): Promise<Finding[]> {
  const scan = await listUnusedStyles().catch(() => null);
  if (scan === null || scan.skippedSheets.length > 0) return [];
  return scan.unused.map((name) =>
    finding("unusedStyle", null, null, name, scan.unused.length),
  );
}

interface SheetScan {
  findings: Finding[];
  sheetOrder: string[];
  scanned: { sheets: number; cells: number };
  skipped: string[];
  brokenNames: string[];
}

function collect(
  sheets: Excel.Worksheet[],
  scanned: ScannedSheet[],
  skipped: string[],
  names: Excel.NamedItemCollection,
): SheetScan {
  return {
    findings: [
      ...hiddenSheetFindings(sheets),
      ...scanned.flatMap(sheetFindings),
    ],
    sheetOrder: sheets.map((sheet) => sheet.name),
    scanned: {
      // Every sheet the pass covered, an empty one included: nothing on it is
      // a complete answer, and only a skipped sheet leaves a hole.
      sheets: sheets.length - skipped.length,
      cells: scanned.reduce((sum, sheet) => sum + sheet.range.cellCount, 0),
    },
    skipped,
    brokenNames: brokenIn(names),
  };
}

// Three syncs, never one per sheet: the sheet list and the names, then every
// used range's extent, then one load of the grids of the sheets small enough to
// read. Hidden sheets are read too - a back room is where the errors hide.
async function scanSheets(
  context: Excel.RequestContext,
  maxCells: number,
  maxTotalCells: number,
): Promise<SheetScan> {
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
    maxCells,
    maxTotalCells,
  );
  for (const sheet of scanned) {
    sheet.range.load("formulas,values,valueTypes,formulasR1C1");
  }
  await context.sync();

  return collect(sheets.items, scanned, skippedSheets, names);
}

/**
 * One pass over the workbook. Read-only: nothing is painted, moved, activated
 * or deleted, and the reviewer decides what to do with every line.
 */
export async function runModelCheck(
  maxCells = MODEL_CHECK_CELL_CAP,
  maxTotalCells = MODEL_CHECK_CELL_CAP,
): Promise<ModelCheckReport> {
  const styles = await unusedStyleFindings();
  const scan = await Excel.run((context) =>
    scanSheets(context, maxCells, maxTotalCells),
  );

  return buildReport({
    findings: [
      ...scan.findings,
      ...scan.brokenNames.map((name) =>
        finding("brokenName", null, null, name),
      ),
      ...styles,
    ],
    sheetOrder: scan.sheetOrder,
    scanned: scan.scanned,
    skipped: scan.skipped,
  });
}
