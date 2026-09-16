// Reading a range as a table: one getCellProperties for the formats, the text
// Excel shows and every column's width, mapped to the cell grid PowerPoint
// rebuilds - and a plain grid of text and widths for a host that refuses the
// formats. Owns the Excel-to-TableCell mapping, what counts as a default
// worth leaving out, and whether row 0 reads as a header. Invariant: a cell
// wearing Excel's own defaults carries nothing but its text, so the deck's
// table style shows through.

import { parseRect } from "../link/geometry";
import { overTableCap, TABLE_TOO_BIG, type TableCell } from "../link/model";
import { localizeNumberText, type NumberSeparators } from "../numbers";
import { originalFillColor } from "./fill-store";
import { hostSupports } from "./internal";
import { bothRefused } from "./link-anchors";
import { parseAddress } from "./shared";

// Excel's defaults. A cell wearing these adds nothing but its size to the payload.
const DEFAULT_FONT_COLOR = "#000000";
// An unfilled cell reads back as white, and a white fill on a white sheet is
// the same picture: both travel as no fill at all.
const NO_FILL = "#FFFFFF";

const ALIGNMENT: Record<string, TableCell["a"]> = {
  Left: "l",
  Center: "c",
  Right: "r",
};

const WANTED: Excel.CellPropertiesLoadOptions = {
  format: {
    font: { bold: true, italic: true, color: true, size: true },
    fill: { color: true },
    horizontalAlignment: true,
  },
};

export interface TableRender {
  rows: number;
  cols: number;
  cells: TableCell[][];
  widths: number[];
}

// Two syncs whatever the size: the grid's dimensions, because the column
// widths have to be asked for one column at a time, then the text, the formats
// and those widths together. A host that refuses the second batch is asked for
// the plain grid instead, in the same context: a table with its values and its
// widths is worth more than an export that stops before the relay, and a later
// push from a host that answers enriches the cells again.
export async function renderTable(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<TableRender> {
  range.load("rowCount,columnCount");
  await context.sync();
  const rows = range.rowCount;
  const cols = range.columnCount;
  if (overTableCap(rows, cols)) throw new Error(TABLE_TOO_BIG);
  try {
    return await renderRichTable(context, range, rows, cols);
  } catch (rich) {
    return await renderPlainTable(context, range, rows, cols, rich);
  }
}

// Queued in the same batch a render already runs, so reading the separators
// costs neither path an extra round trip: null on a host below ExcelApi
// 1.11, the guard every adapter in this file already uses before it reads
// them (readSeparators resolves the value once that sync commits).
function queueSeparators(
  context: Excel.RequestContext,
): Excel.Application | null {
  if (!hostSupports("1.11")) return null;
  const application = context.application;
  application.load("decimalSeparator,thousandsSeparator");
  return application;
}

function readSeparators(
  application: Excel.Application | null,
): NumberSeparators | null {
  return application
    ? {
        decimal: application.decimalSeparator,
        thousands: application.thousandsSeparator,
      }
    : null;
}

// Sync two: the formats, the text, every column's width and the application's
// separators in one batch.
async function renderRichTable(
  context: Excel.RequestContext,
  range: Excel.Range,
  rows: number,
  cols: number,
): Promise<TableRender> {
  const properties = range.getCellProperties(WANTED);
  range.load("text,values,address,worksheet/id");
  const columns = Array.from({ length: cols }, (_unused, index) => {
    const column = range.getColumn(index);
    column.load("format/columnWidth");
    return column;
  });
  const application = queueSeparators(context);
  await context.sync();

  const text = range.text;
  const values = range.values;
  const separators = readSeparators(application);
  const sheetId = range.worksheet.id;
  // Null when the address is not one geometry.ts can read (a multi-area
  // range, say): no cell of this render can be inside a store's snapshot
  // then, so every fill falls back to the live read.
  const origin = parseRect(parseAddress(range.address).address);
  const cells = properties.value.map((row, r) =>
    row.map((cell, c) => {
      const fillOverride = origin
        ? originalFillColor(sheetId, origin.top + r, origin.left + c)
        : undefined;
      return toCell(
        text[r]?.[c] ?? "",
        cell,
        typeof values[r]?.[c] === "number",
        separators,
        fillOverride,
      );
    }),
  );
  const widths = columns.map((one) => one.format.columnWidth);
  return { rows, cols, cells, widths };
}

// What the deck can still be given when the formats are refused: the text
// Excel shows, a number's own right alignment and the column widths. The reads
// are queued fresh because the batch that asked for them was never committed.
async function renderPlainTable(
  context: Excel.RequestContext,
  range: Excel.Range,
  rows: number,
  cols: number,
  rich: unknown,
): Promise<TableRender> {
  try {
    range.load("text,values");
    const columns = Array.from({ length: cols }, (_unused, index) => {
      const column = range.getColumn(index);
      column.load("format/columnWidth");
      return column;
    });
    const application = queueSeparators(context);
    await context.sync();
    const separators = readSeparators(application);
    const cells = range.text.map((row, r) =>
      row.map((text, c) => {
        const numeric = typeof range.values[r]?.[c] === "number";
        const localized =
          numeric && separators ? localizeNumberText(text, separators) : text;
        return {
          t: localized,
          ...(numeric ? { a: "r" as const } : {}),
        };
      }),
    );
    return {
      rows,
      cols,
      cells,
      widths: columns.map((one) => one.format.columnWidth),
    };
  } catch (plain) {
    throw bothRefused(rich, "rich", plain, "plain");
  }
}

// Every key but the text is omitted unless the cell says something Excel's
// default does not, which is what keeps a plain grid small.
// A number under Excel's "General" alignment sits on the right; PowerPoint
// has no such rule, so the cell says so itself.
//
// `fillOverride` is undefined when no FillStore owns this cell (the live
// read is the true fill), null when one owns it and remembers no fill, or
// the colour it remembers: an overlay's tint never ships in a cell's place.
function toCell(
  text: string,
  properties: Excel.CellProperties,
  numeric: boolean,
  separators: NumberSeparators | null,
  fillOverride: string | null | undefined,
): TableCell {
  const format = properties.format;
  const font = format?.font;
  // Office.js always hands back a number's text with the invariant comma and
  // point; a text cell, or one under separators Excel already shows, is
  // untouched by localizeNumberText itself.
  const localized =
    numeric && separators ? localizeNumberText(text, separators) : text;
  const cell: TableCell = { t: localized };
  if (font?.bold === true) cell.b = true;
  if (font?.italic === true) cell.i = true;
  if (font?.color !== undefined && font.color !== DEFAULT_FONT_COLOR) {
    cell.c = font.color;
  }
  // The size always travels: a deck's table style defaults to 18 pt, a
  // model's cells to 11, and a table twice the size of its source is no copy.
  if (font?.size !== undefined) cell.z = font.size;
  const fill =
    fillOverride === undefined
      ? format?.fill?.color
      : (fillOverride ?? undefined);
  if (fill !== undefined && fill !== NO_FILL && fill !== "") cell.f = fill;
  const alignment = ALIGNMENT[String(format?.horizontalAlignment)];
  if (alignment !== undefined) cell.a = alignment;
  else if (numeric) cell.a = "r";
  return cell;
}

// A table's first row reads as a header when every one of its non-empty
// cells came out bold: a bold label over plain numbers is not a header, and
// a single-row export has no body for a header to sit above.
export function headerRow(cells: TableCell[][]): boolean {
  if (cells.length < 2) return false;
  const labelled = (cells[0] ?? []).filter((cell) => cell.t.trim() !== "");
  return labelled.length > 0 && labelled.every((cell) => cell.b === true);
}
