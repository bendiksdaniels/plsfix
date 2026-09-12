// Workbook tools: the contents (TOC) sheet, the sheet explorer (list/show/hide/
// activate), the sheet tools (unhide all, show only this, bury, move) and the
// broken-name scrubber. The TOC sheet is rewritten in full on every run and
// only ever touches a sheet it marked as its own (the A1 marker).
//
// The sheet tools run outside pls,fix Undo, the way the size cycles in
// sizes.ts do: visibility and position are sheet state, getCellProperties
// carries neither, so captureUndo could not put one back and is not called.
// Every one of them ends on syncWrite with structureNote, so a workbook whose
// structure is protected answers the pane's sentence, not Excel's own string.

import {
  brokenEverywhere,
  brokenIn,
  loadNames,
  loadSheetNames,
} from "./internal";
import { structureNote, structureProtected, syncWrite } from "./protection";
import { activeTheme, getActiveSettings } from "../settings";
import { tocRows } from "../workbook";

const TOC_SHEET = "TOC";
const TOC_MARKER = "pls,fix - Contents";
const TOC_INDEX_WIDTH = 34;
const TOC_NAME_WIDTH = 240;

// Rewritten in full on every run, so it can never drift from the workbook. That
// only holds for a sheet we wrote ourselves: A1 carries the marker that says so.
export async function insertToc(): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const existing = sheets.getItemOrNullObject(TOC_SHEET);
    existing.load("isNullObject");
    sheets.load("items/name,items/visibility");
    await context.sync();

    // Snapshot the sheet list before adding anything: the contents sheet does
    // not list itself.
    const listed = sheets.items
      .map((item) => ({ name: item.name, visibility: item.visibility }))
      .filter((item) => item.name.toLowerCase() !== TOC_SHEET.toLowerCase());

    let sheet = existing;
    if (existing.isNullObject) {
      sheet = sheets.add(TOC_SHEET);
      sheet.position = 0;
    } else {
      const marker = existing.getRange("A1");
      marker.load("values");
      const used = existing.getUsedRangeOrNullObject();
      used.load("isNullObject");
      await context.sync();

      if (marker.values[0]?.[0] !== TOC_MARKER) {
        throw new Error("A sheet named TOC already exists and is not ours.");
      }
      if (!used.isNullObject) {
        used.clear(Excel.ClearApplyTo.all);
        // A shorter list must not leave a live link behind in a cleared cell.
        used.clear(Excel.ClearApplyTo.removeHyperlinks);
      }
    }

    const rows = tocRows(listed);
    sheet.getRange("A1").values = [[TOC_MARKER]];
    rows.forEach((row, offset) => {
      sheet.getRangeByIndexes(2 + offset, 0, 1, 1).values = [[row.index]];
      sheet.getRangeByIndexes(2 + offset, 1, 1, 1).hyperlink = {
        documentReference: row.target,
        textToDisplay: row.name,
      };
    });
    // Excel dresses a new hyperlink in its own blue style, so the brand colors
    // go on afterwards rather than in the same batch.
    await context.sync();

    const theme = activeTheme();
    const { font } = getActiveSettings();
    const title = sheet.getRange("A1:B1").format;
    title.font.name = font;
    title.font.size = 15;
    title.font.bold = true;
    title.font.color = theme.titleText;
    title.fill.color = theme.titleFill;
    title.verticalAlignment = Excel.VerticalAlignment.center;
    title.rowHeight = 25;

    if (rows.length > 0) {
      const body = sheet.getRangeByIndexes(2, 0, rows.length, 2).format;
      body.font.name = font;
      body.font.size = 10;
      body.font.bold = false;
      body.font.color = theme.formulaFont;
      body.verticalAlignment = Excel.VerticalAlignment.center;

      const links = sheet.getRangeByIndexes(2, 1, rows.length, 1).format;
      links.font.color = theme.linkFont;
      links.font.underline = Excel.RangeUnderlineStyle.none;
      sheet.getRangeByIndexes(2, 0, rows.length, 1).format.horizontalAlignment =
        Excel.HorizontalAlignment.right;
    }

    sheet.getRange("A:A").format.columnWidth = TOC_INDEX_WIDTH;
    sheet.getRange("B:B").format.columnWidth = TOC_NAME_WIDTH;
    sheet.showGridlines = false;
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.activate();
    await context.sync();
  });
}

export interface SheetEntry {
  name: string;
  visibility: string;
  active: boolean;
}

export async function listSheets(): Promise<SheetEntry[]> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility,items/id");
    const active = sheets.getActiveWorksheet();
    active.load("id");
    await context.sync();

    return sheets.items.map((item) => ({
      name: item.name,
      visibility: item.visibility,
      active: item.id === active.id,
    }));
  });
}

export async function setSheetVisibility(
  name: string,
  visible: boolean,
): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItem(name);
    sheets.load("items/visibility");
    sheet.load("visibility");
    await context.sync();

    // VeryHidden is a deliberate authoring choice made outside Excel's UI; the
    // explorer reports it and leaves it alone.
    if (sheet.visibility === Excel.SheetVisibility.veryHidden) {
      throw new Error(`${name} is very hidden and can only be shown in VBA.`);
    }
    const visibleCount = sheets.items.filter(
      (item) => item.visibility === Excel.SheetVisibility.visible,
    ).length;
    if (!visible && visibleCount <= 1) {
      throw new Error("A workbook needs at least one visible sheet.");
    }

    sheet.visibility = visible
      ? Excel.SheetVisibility.visible
      : Excel.SheetVisibility.hidden;
    await context.sync();
  });
}

export async function activateSheet(name: string): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItem(name).activate();
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Sheet tools: the six buttons above the explorer. Each one asks first, so a
// locked sheet list is refused before anything is written rather than after.
// ---------------------------------------------------------------------------

export type SheetMove = "up" | "down" | "end";

/** Where a move lands the sheet. Already there is not a move (see moveSheet). */
function movedTo(direction: SheetMove, from: number, last: number): number {
  if (direction === "up") return Math.max(0, from - 1);
  if (direction === "down") return Math.min(last, from + 1);
  return last;
}

async function requireSheetList(
  context: Excel.RequestContext,
  stage: string,
): Promise<void> {
  if (await structureProtected(context)) throw new Error(structureNote(stage));
}

/** What one unhide-all pass found: what it showed, and what it left buried. */
export interface UnhideResult {
  shown: number;
  buried: number;
}

/**
 * Unhide all: every hidden sheet back on show. A very hidden sheet was put out
 * of sight outside Excel's UI, so it only comes back when asked for; the ones
 * left behind are counted, because "no hidden sheets" would otherwise be a lie
 * on a workbook whose only hidden sheets are buried.
 */
export async function setSheetsVisibility(
  includeVeryHidden: boolean,
): Promise<UnhideResult> {
  return Excel.run(async (context) => {
    await requireSheetList(context, "Unhide all");
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();

    const veryHidden = sheets.items.filter(
      (item) => item.visibility === Excel.SheetVisibility.veryHidden,
    );
    const hidden = sheets.items.filter(
      (item) =>
        item.visibility === Excel.SheetVisibility.hidden ||
        (includeVeryHidden &&
          item.visibility === Excel.SheetVisibility.veryHidden),
    );
    for (const item of hidden) {
      item.visibility = Excel.SheetVisibility.visible;
    }
    await syncWrite(context, "Unhide all", structureNote);
    return {
      shown: hidden.length,
      buried: includeVeryHidden ? 0 : veryHidden.length,
    };
  });
}

/** Show only this: every other visible sheet hidden, this one on show. */
export async function showOnlySheet(): Promise<{
  name: string;
  hidden: number;
}> {
  return Excel.run(async (context) => {
    await requireSheetList(context, "Show only this");
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/id,items/visibility");
    active.load("id,name");
    await context.sync();

    const others = sheets.items.filter(
      (item) =>
        item.id !== active.id &&
        item.visibility === Excel.SheetVisibility.visible,
    );
    for (const item of others) item.visibility = Excel.SheetVisibility.hidden;
    active.visibility = Excel.SheetVisibility.visible;
    await syncWrite(context, "Show only this", structureNote);
    return { name: active.name, hidden: others.length };
  });
}

/** Bury this: the active sheet out of the tab strip and out of the unhide list. */
export async function burySheet(): Promise<string> {
  return Excel.run(async (context) => {
    await requireSheetList(context, "Bury this sheet");
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/visibility");
    active.load("name,visibility");
    await context.sync();

    const visible = sheets.items.filter(
      (item) => item.visibility === Excel.SheetVisibility.visible,
    ).length;
    if (active.visibility === Excel.SheetVisibility.visible && visible <= 1) {
      throw new Error("Excel needs one visible sheet.");
    }

    active.visibility = Excel.SheetVisibility.veryHidden;
    await syncWrite(context, "Bury this sheet", structureNote);
    return active.name;
  });
}

/** Move the active sheet one place, or to the end of the tab strip. */
export async function moveSheet(direction: SheetMove): Promise<{
  name: string;
  position: number;
}> {
  return Excel.run(async (context) => {
    await requireSheetList(context, "Move sheet");
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/name");
    active.load("name,position");
    await context.sync();

    const from = active.position;
    const to = movedTo(direction, from, sheets.items.length - 1);
    if (to === from) {
      const edge = direction === "up" ? "first" : "last";
      throw new Error(`${active.name} is already the ${edge} sheet.`);
    }

    active.position = to;
    await syncWrite(context, "Move sheet", structureNote);
    return { name: active.name, position: to };
  });
}

// A name can be scoped to the workbook or to one sheet (Worksheet.names,
// ExcelApi 1.4): the two collections are separate, so a broken name defined on
// a sheet never showed up in workbook.names at all. loadSheetNames (in
// internal.ts, shared with share.ts and model-check.ts) loads the second sheet
// by sheet, once the sheet list itself is in hand.
async function loadEveryNameCollection(context: Excel.RequestContext): Promise<{
  workbook: Excel.NamedItemCollection;
  sheets: Excel.Worksheet[];
  perSheet: Excel.NamedItemCollection[];
}> {
  const workbook = loadNames(context);
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const perSheet = loadSheetNames(sheets.items);
  await context.sync();

  return { workbook, sheets: sheets.items, perSheet };
}

export async function listBrokenNames(): Promise<string[]> {
  return Excel.run(async (context) => {
    const { workbook, sheets, perSheet } =
      await loadEveryNameCollection(context);
    return brokenEverywhere(workbook, sheets, perSheet);
  });
}

const STRUCTURE_PROTECTED =
  "The workbook's structure is protected, so nothing was deleted.";

// Irreversible: pls,fix Undo restores ranges, and Office.js writes never reach
// Excel's own undo stack. The pane therefore asks twice before calling this.
export async function deleteBrokenNames(): Promise<number> {
  return Excel.run(async (context) => {
    const { workbook, perSheet } = await loadEveryNameCollection(context);

    const broken = new Set(brokenIn(workbook));
    for (const item of workbook.items) {
      if (broken.has(item.name)) item.delete();
    }

    let deleted = broken.size;
    for (const collection of perSheet) {
      const scopedBroken = new Set(brokenIn(collection));
      for (const item of collection.items) {
        if (scopedBroken.has(item.name)) {
          item.delete();
          deleted += 1;
        }
      }
    }

    try {
      await context.sync();
    } catch (error) {
      const { code } = error as { code?: string };
      if (code === Excel.ErrorCodes.accessDenied) {
        throw new Error(STRUCTURE_PROTECTED);
      }
      throw error;
    }
    return deleted;
  });
}
