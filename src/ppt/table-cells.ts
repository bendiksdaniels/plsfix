// Writing a table's cells a few at a time: which cells a format pass needs to
// touch, chunked so no round trip is big enough to reach the sync deadline,
// and the per-cell property writes themselves. Split out of tables.ts, which
// sits at the file-length cap. Invariant: a repaint writes every property of
// a cell it touches; an insert or a rebuild only ever adds one.

import { type TableCell, type TablePayload } from "../link/model";
import { paintKey } from "../link/paint-map";
import { withSyncDeadline } from "./chart-draw";

// How many cells one round trip formats. PowerPoint for the web spent about
// 0.4 s per cell property write on 09.09, and a repaint writes five per cell,
// so eight cells keep a sync near 15 s, well inside SYNC_TIMEOUT_MS, where a
// whole 6x4 table in one batch ran past the deadline on the rig.
export const CELLS_PER_SYNC = 8;
const ALIGNMENT = { l: "Left", c: "Center", r: "Right" } as const;

export const FORMATTING = "formatting the table";
export const REPAINTING = "repainting the table";
// Insert and rebuild only ever add a fill; nothing painted before exists yet
// for either to undo.
export const NO_CLEARS: ReadonlySet<string> = new Set();

type CellAt = [row: number, column: number, cell: TableCell];

// Text is written on a repaint only: an insert carries it in `values`, so a
// plain cell then costs no call at all - which is what keeps a plain 60x20
// table one round trip rather than twelve hundred.
function cellChunks(payload: TablePayload, withText: boolean): CellAt[][] {
  const cells: CellAt[] = [];
  payload.cells.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (withText || formatted(cell))
        cells.push([rowIndex, columnIndex, cell]);
    });
  });
  const chunks: CellAt[][] = [];
  for (let start = 0; start < cells.length; start += CELLS_PER_SYNC) {
    chunks.push(cells.slice(start, start + CELLS_PER_SYNC));
  }
  return chunks;
}

// CELLS_PER_SYNC cells per round trip, so no batch is big enough to reach the
// deadline and one the host swallows loses a chunk rather than the whole
// table; `clear` names the cells (paintKey) allowed to lose their fill, and
// `beforeLast` queues a tag write for the table's final round trip.
// `uniformFontSize` is the size addTable already gave every cell (insert and
// rebuild only: a repaint's table was built at some earlier revision's size,
// never this one, so it always writes cell.z and passes none).
interface WriteOptions {
  withText: boolean;
  clear: ReadonlySet<string>;
  what: string;
  uniformFontSize?: number;
  beforeLast?: () => void;
}

export async function writeCellsInChunks(
  context: PowerPoint.RequestContext,
  table: PowerPoint.Table,
  payload: TablePayload,
  options: WriteOptions,
): Promise<void> {
  const { withText, clear, what, uniformFontSize, beforeLast } = options;
  const chunks = cellChunks(payload, withText);
  if (chunks.length === 0) {
    if (!beforeLast) return;
    beforeLast();
    await withSyncDeadline(context.sync(), what);
    return;
  }
  for (const [index, chunk] of chunks.entries()) {
    for (const [row, column, cell] of chunk) {
      writeCell(
        table.getCellOrNullObject(row, column),
        cell,
        withText,
        clear.has(paintKey(row, column)),
        uniformFontSize,
      );
    }
    if (index === chunks.length - 1) beforeLast?.();
    await withSyncDeadline(context.sync(), what);
  }
}

function formatted(cell: TableCell): boolean {
  return [cell.b, cell.i, cell.c, cell.f, cell.a, cell.z].some(
    (value) => value !== undefined,
  );
}

function writeCell(
  target: PowerPoint.TableCell,
  cell: TableCell,
  withText: boolean,
  clear: boolean,
  uniformFontSize?: number,
): void {
  if (withText) target.text = cell.t;
  // A repaint has to undo what the last one wrote; an insert only ever adds.
  if (withText || cell.b !== undefined) target.font.bold = cell.b === true;
  if (withText || cell.i !== undefined) target.font.italic = cell.i === true;
  if (withText || cell.a !== undefined) {
    target.horizontalAlignment = ALIGNMENT[cell.a ?? "l"];
  }
  if (cell.f !== undefined) target.fill.setSolidColor(cell.f);
  else if (clear) target.fill.clear();
  // Colour and size have no "back to the table style" to write, so a cell that
  // stops naming them keeps what it had until the table is built again.
  if (cell.c !== undefined) target.font.color = cell.c;
  // A size equal to what addTable's uniformCellProperties already gave the
  // cell is one property write saved - about 0.4 s on PowerPoint for the web;
  // a cell whose own size differs still gets it written.
  if (cell.z !== undefined && cell.z !== uniformFontSize) {
    target.font.size = cell.z;
  }
}
