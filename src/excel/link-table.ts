// Reading a range as a table: one getCellProperties for the formats, the text
// Excel shows and every column's width, mapped to the cell grid PowerPoint
// rebuilds. Owns the Excel-to-TableCell mapping and what counts as a default
// worth leaving out. Invariant: a cell wearing Excel's own defaults carries
// nothing but its text, so the deck's table style shows through.

import { overTableCap, TABLE_TOO_BIG, type TableCell } from "../link/model";
import { isMacExcel } from "./link-platform";

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
// and those widths together.
export async function renderTable(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<TableRender> {
  range.load("rowCount,columnCount");
  await context.sync();
  const rows = range.rowCount;
  const cols = range.columnCount;
  if (overTableCap(rows, cols)) throw new Error(TABLE_TOO_BIG);

  // Mac's rich CellProperties grid is the other export call known to fail
  // before a relay request. Preserve the editable table's displayed values,
  // numeric alignment and column sizing instead of failing the whole export;
  // chart/table source data remains intact and a later push on another host
  // can enrich its per-cell formatting again.
  if (isMacExcel()) return renderMacTable(context, range, rows, cols);

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

async function renderMacTable(
  context: Excel.RequestContext,
  range: Excel.Range,
  rows: number,
  cols: number,
): Promise<TableRender> {
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
      ...(typeof range.values[r]?.[c] === "number" ? { a: "r" as const } : {}),
    })),
  );
  return {
    rows,
    cols,
    cells,
    widths: columns.map((one) => one.format.columnWidth),
  };
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
