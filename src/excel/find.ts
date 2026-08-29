// Super Find: one pass over the whole workbook - every sheet's used range, the
// defined names, the sheet names themselves and the comments - and the jump
// that follows a result back to its cell. Four syncs, one per phase, never one
// per sheet or per comment: the sheet list, the names and the comments, then
// every used range's extent, then the grids of the sheets small enough to read,
// then every comment's cell and reply thread in one batch (and that fourth sync
// only where there are comments at all).

import {
  cellAddress,
  COMMENT_ROW,
  type CommentEntry,
  includesQuery,
  matchCells,
  matchComments,
  type MatchOptions,
  type RankedHit,
  NAME_ORDER,
  rankHits,
  SHEET_NAME_COL,
  SHEET_NAME_ROW,
} from "../find";
import {
  hostSupports,
  pickScannableSheets,
  SCAN_CELL_CAP,
  type ScannedSheet,
  SELECTION_CELL_CAP,
} from "./internal";
import { ANCHOR_PREFIX } from "../link/model";
import { type CellValue } from "../model";
import { parseAddress } from "./shared";

// Comments arrived with ExcelApi 1.10; older hosts have no collection to read.
const COMMENT_API_SET = "1.10";

// What every comment says and who wrote it. A reply thread is not a load option
// of the comment collection, so it cannot be expanded from here: the replies are
// asked for one level down, beside the locations.
const COMMENT_LOAD = "items/content,items/authorName";
const REPLY_LOAD = "items/content,items/authorName";

export interface FindHit {
  kind: "cell" | "name" | "sheet" | "comment";
  sheet: string;
  address: string;
  text: string;
}

export interface FindOptions extends MatchOptions {
  // The used range a sheet may have before it is skipped instead of read.
  // Defaults to the selection cap; the tests use a smaller one.
  maxCells?: number;
  // What the whole scan may read across every sheet, so a workbook of middling
  // sheets cannot queue one request the host refuses.
  maxTotalCells?: number;
  // Comments are their own phase with their own host requirement, so the pane
  // can leave them out; on unless the box is unticked, as it is on screen.
  inComments?: boolean;
}

export interface FindResult {
  hits: FindHit[];
  // Sheets whose used range was too large to read, named so the pane can say
  // the answer is incomplete rather than quietly leaving them out.
  skippedSheets: string[];
  // The search wanted comments and the host is below ExcelApi 1.10: the pane
  // says so rather than letting a modeller read "no matches" as "none there".
  commentsSkipped: boolean;
}

// A comment and each of its replies become one row apiece; only the adapter
// knows which is which, so a query never matches the "(reply)" mark itself.
interface CommentRecord extends CommentEntry {
  reply: boolean;
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
      col: SHEET_NAME_COL,
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

// The box is ticked by default, so an option that says nothing means yes.
function wantsComments(options: FindOptions): boolean {
  return options.inComments !== false;
}

// The comments are asked for in the first batch, beside the names: their cells
// and threads cannot be, because both need the items themselves in hand.
function loadComments(
  context: Excel.RequestContext,
  options: FindOptions,
): Excel.CommentCollection | null {
  if (!wantsComments(options) || !hostSupports(COMMENT_API_SET)) return null;
  const comments = context.workbook.comments;
  comments.load(COMMENT_LOAD);
  return comments;
}

// One flat row per comment and per reply, all on the comment's own cell.
function commentRecords(
  comments: Excel.CommentCollection,
  locations: Excel.Range[],
): CommentRecord[] {
  const records: CommentRecord[] = [];
  comments.items.forEach((comment, index) => {
    const at = parseAddress(locations[index]?.address ?? "");
    records.push({
      ...at,
      content: comment.content,
      author: comment.authorName,
      reply: false,
    });
    for (const reply of comment.replies.items) {
      records.push({
        ...at,
        content: reply.content,
        author: reply.authorName,
        reply: true,
      });
    }
  });
  return records;
}

function commentRows(
  records: CommentRecord[],
  sheetOrder: Map<string, number>,
  query: string,
  options: FindOptions,
): ScanRow[] {
  const rows: ScanRow[] = [];
  for (const hit of matchComments(records, query, options)) {
    const record = records[hit.order];
    if (!record) continue;
    const sheetIndex = sheetOrder.get(record.sheet);
    if (sheetIndex === undefined) continue;
    rows.push({
      sheetIndex,
      row: COMMENT_ROW,
      col: hit.order,
      hit: {
        kind: "comment",
        sheet: record.sheet,
        address: record.address,
        text: record.reply ? `(reply) ${hit.text}` : hit.text,
      },
    });
  }
  return rows;
}

// Phase 4, and the one extra sync: every comment's cell and every reply thread
// are queued in the same batch, so a workbook with comments costs one round
// trip more however many of them there are, and one without costs nothing.
async function commentHits(
  context: Excel.RequestContext,
  comments: Excel.CommentCollection | null,
  sheets: Excel.Worksheet[],
  query: string,
  options: FindOptions,
): Promise<ScanRow[]> {
  if (!comments || comments.items.length === 0) return [];
  const locations = comments.items.map((comment) => {
    comment.replies.load(REPLY_LOAD);
    const location = comment.getLocation();
    location.load("address");
    return location;
  });
  await context.sync();

  const order = new Map<string, number>(
    sheets.map((sheet, index) => [sheet.name, index]),
  );
  return commentRows(
    commentRecords(comments, locations),
    order,
    query,
    options,
  );
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
    const comments = loadComments(context, options);
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
    for (const sheet of scanned) sheet.range.load("values,formulas");
    await context.sync();

    const rows = [
      ...nameHits(names.items, query, options),
      ...sheetHits(sheets.items, query, options),
      ...cellHits(scanned, query, options),
      ...(await commentHits(context, comments, sheets.items, query, options)),
    ];
    return {
      hits: rankHits(rows).map((row) => row.hit),
      skippedSheets,
      commentsSkipped: wantsComments(options) && comments === null,
    };
  });
}

// A name is matched on its formula as well as on its own name, so a hit can be
// a constant - "Tax" finds TaxRate = 0.21 - or a name Excel has rewritten to
// #REF!. Neither points at a range, and getRange() answers that with a bare
// host string, so the null object is asked for instead and the refusal names
// the flow and the name.
async function namedRange(
  context: Excel.RequestContext,
  name: string,
): Promise<Excel.Range> {
  const range = context.workbook.names.getItem(name).getRangeOrNullObject();
  range.load("isNullObject");
  await context.sync();
  if (range.isNullObject) {
    throw new Error(`find: name "${name}" has no range`);
  }
  return range;
}

// Excel cannot select on a sheet it is not showing, so the jump activates first
// and refuses a hidden sheet by name rather than letting the host throw. A
// comment jumps to the cell it hangs on, exactly as a cell hit does.
export async function jumpToHit(hit: FindHit): Promise<void> {
  await Excel.run(async (context) => {
    const range =
      hit.kind === "name"
        ? await namedRange(context, hit.address)
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
