// Workbook tools: the contents (TOC) sheet, the sheet explorer (list/show/hide/
// activate) and the broken-name scrubber. The TOC sheet is rewritten in full on
// every run and only ever touches a sheet it marked as its own (the A1 marker).

import { brokenIn, loadNames } from "./internal";
import { activeTheme, getActiveSettings } from "../settings";
import { tocRows } from "../workbook";

const TOC_SHEET = "TOC";
const TOC_MARKER = "Model Tools - Contents";
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

export async function listBrokenNames(): Promise<string[]> {
  return Excel.run(async (context) => {
    const names = loadNames(context);
    await context.sync();
    return brokenIn(names);
  });
}

// Irreversible: SMT Undo restores ranges, and Office.js writes never reach
// Excel's own undo stack. The pane therefore asks twice before calling this.
export async function deleteBrokenNames(): Promise<number> {
  return Excel.run(async (context) => {
    const names = loadNames(context);
    await context.sync();

    const broken = new Set(brokenIn(names));
    for (const item of names.items) {
      if (broken.has(item.name)) item.delete();
    }
    await context.sync();
    return broken.size;
  });
}
