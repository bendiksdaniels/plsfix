// Reading a range as a table: one getCellProperties for the formats, the text
// Excel shows and every column's width, mapped to the cell grid PowerPoint
// rebuilds - and a plain grid of text and widths for a host that refuses the
// formats. Owns the Excel-to-TableCell mapping and what counts as a default
// worth leaving out. Invariant: a cell wearing Excel's own defaults carries
// nothing but its text, so the deck's table style shows through.

import { overTableCap, TABLE_TOO_BIG, type TableCell } from "../link/model";
import { bothRefused } from "./link-anchors";

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

// Sync two: the formats, the text and every column's width in one batch.
async function renderRichTable(
  context: Excel.RequestContext,
  range: Excel.Range,
  rows: number,
  cols: number,
): Promise<TableRender> {
  const properties = range.getCellProperties(WANTED);
  range.load("text,values");
  const columns = Array.from({ length: cols }, (_unused, index) => {
    const column = range.getColumn(index);
    column.load("format/columnWidth");
    return column;
  });
  await context.sync();

  const text = range.text;
  const values = range.values;
  const cells = properties.value.map((row, r) =>
    row.map((cell, c) =>
      toCell(text[r]?.[c] ?? "", cell, typeof values[r]?.[c] === "number"),
    ),
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
    await context.sync();
    const cells = range.text.map((row, r) =>
      row.map((text, c) => ({
        t: text,
        ...(typeof range.values[r]?.[c] === "number"
          ? { a: "r" as const }
          : {}),
      })),
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
function toCell(
  text: string,
  properties: Excel.CellProperties,
  numeric: boolean,
): TableCell {
  const format = properties.format;
  const font = format?.font;
  const cell: TableCell = { t: text };
  if (font?.bold === true) cell.b = true;
  if (font?.italic === true) cell.i = true;
  if (font?.color !== undefined && font.color !== DEFAULT_FONT_COLOR) {
    cell.c = font.color;
  }
  // The size always travels: a deck's table style defaults to 18 pt, a
  // model's cells to 11, and a table twice the size of its source is no copy.
  if (font?.size !== undefined) cell.z = font.size;
  const fill = format?.fill?.color;
  if (fill !== undefined && fill !== NO_FILL && fill !== "") cell.f = fill;
  const alignment = ALIGNMENT[String(format?.horizontalAlignment)];
  if (alignment !== undefined) cell.a = alignment;
  else if (numeric) cell.a = "r";
  return cell;
}
