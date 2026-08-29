// Prepare for sharing: the tidy pass a model gets before it leaves the desk.
// Every visible sheet goes back to A1 and the workbook lands on the first of
// them, and one scan reports what a reader would still find - hidden sheets,
// formulas pointing at other workbooks or needing this add-in, names left on
// #REF!, an overlay still painted over the model, the link tokens the file
// carries and autocolor still recoloring every edit. Nothing is deleted, no
// hidden sheet is touched, and no link is broken: the report says so and the
// modeller decides.
// Zoom is not in the Office.js worksheet surface (pageLayout.zoom is print
// zoom), so the pane says so rather than pretending to reset it.

import { auditOverlayOn } from "./audit";
import { autocolorOnEditActive } from "./autocolor";
import {
  brokenIn,
  loadNames,
  pickScannableSheets,
  SCAN_CELL_CAP,
  type ScannedSheet,
  SELECTION_CELL_CAP,
} from "./internal";
import { linkHighlightOn } from "./link-highlight";
import { cellAddress } from "../find";
import { REGISTRY_SETTING, tryDecodeRegistry } from "../link/model";
import { type CellValue } from "../model";
import {
  isAddinFormula,
  isExternalFormula,
  type ShareIssue,
  type ShareLink,
  shareReport,
} from "../share";

export interface ShareOptions {
  // The used range a sheet may have before it is skipped instead of read.
  // Defaults to the selection cap; the tests use a smaller one.
  maxCells?: number;
  // What the whole scan may read across every sheet.
  maxTotalCells?: number;
}

export interface ShareResult {
  report: ShareIssue[];
  // Visible sheets put back at A1, which is what the summary line counts.
  touchedSheets: number;
}

// One walk over the formulas, two answers: what points at another workbook and
// what needs this add-in to evaluate at all. The grid is read from the used
// range, so a hit's coordinates are relative to its top-left corner; the
// address the reader is given is the absolute one.
function scanFormulas(sheets: ScannedSheet[]): {
  external: ShareLink[];
  addin: ShareLink[];
} {
  const external: ShareLink[] = [];
  const addin: ShareLink[] = [];
  for (const sheet of sheets) {
    const formulas = sheet.range.formulas as CellValue[][];
    formulas.forEach((row, rowOffset) => {
      row.forEach((formula, colOffset) => {
        const external_ = isExternalFormula(formula);
        if (!external_ && !isAddinFormula(formula)) return;
        const link: ShareLink = {
          sheet: sheet.name,
          address: cellAddress(
            sheet.range.rowIndex + rowOffset,
            sheet.range.columnIndex + colOffset,
          ),
          formula: String(formula),
        };
        (external_ ? external : addin).push(link);
      });
    });
  }
  return { external, addin };
}

// The overlays holding fills right now, by the name the pane calls them. Both
// paint into the file, and a reader without the add-in has no way to take
// either off.
function paintedOverlays(): string[] {
  return [
    ...(auditOverlayOn() ? ["The audit overlay"] : []),
    ...(linkHighlightOn() ? ["The linked-cell highlight"] : []),
  ];
}

// Counted, never rewritten: an entry a newer build wrote still means the
// workbook carries tokens, whether this one can decode it or not.
function carriesLinkTokens(setting: Excel.Setting): boolean {
  if (setting.isNullObject) return false;
  const raw = String(setting.value);
  const registry = tryDecodeRegistry(raw);
  return registry === null ? raw.trim() !== "" : registry.links.length > 0;
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

// Everything the report is made of, once the three reading phases are done.
function buildReport(
  sheets: Excel.Worksheet[],
  scanned: ScannedSheet[],
  names: Excel.NamedItemCollection,
  skippedSheets: string[],
  registry: Excel.Setting,
): ShareIssue[] {
  const formulas = scanFormulas(scanned);
  return shareReport({
    sheets: sheets.map((sheet) => ({
      name: sheet.name,
      visibility: sheet.visibility,
    })),
    externalLinks: formulas.external,
    addinFormulas: formulas.addin,
    brokenNames: brokenIn(names),
    skippedSheets,
    overlaysPainted: paintedOverlays(),
    linkTokens: carriesLinkTokens(registry),
    autocolorOnEdit: autocolorOnEditActive(),
  });
}

// Four syncs, one per phase, never one per sheet: the sheet list, the names and
// the link registry, then every used range's extent, then the formulas of the
// sheets small enough to read, then the whole A1 reset in a single batch.
export async function prepareForSharing(
  options: ShareOptions = {},
): Promise<ShareResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    const names = loadNames(context);
    // In the same batch: the link registry lives in the file, so what it holds
    // is part of what a reader would receive.
    const registry =
      context.workbook.settings.getItemOrNullObject(REGISTRY_SETTING);
    registry.load("isNullObject,value");
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
      options.maxTotalCells ?? SCAN_CELL_CAP,
    );
    for (const sheet of scanned) sheet.range.load("formulas");
    await context.sync();

    const report = buildReport(
      sheets.items,
      scanned,
      names,
      skippedSheets,
      registry,
    );
    const touchedSheets = resetToA1(sheets.items);
    await context.sync();
    return { report, touchedSheets };
  });
}
