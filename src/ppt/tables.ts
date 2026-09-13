// Native PowerPoint tables (PowerPointApi 1.8): insert a table link at a free
// spot on the slide, write every cell's text and formats, and repaint one in
// place while its dimensions still match - building it again at the same
// corner when they do not. Owns every table-shaped Office.js call.
// Invariant: a repaint writes cells, never geometry.

import type { Box, Size } from "../layout";
import {
  encodeTag,
  overTableCap,
  sourceLabel,
  TABLE_TOO_BIG,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type TableCell,
  type TablePayload,
} from "../link/model";
import { cleanupShapes } from "./chart-cleanup";
import { withSyncDeadline } from "./chart-draw";
import type { FoundLink, InsertResult } from "./host";
import {
  CONTENT_WIDTH,
  DEFAULT_TARGET,
  finishTarget,
  resolveTarget,
  type InsertTarget,
} from "./placement";
import { hasPowerPointApi, isGrouped, shapeAt } from "./shapes";

// shapes.addTable and the Table object both arrived in PowerPointApi 1.8.
const TABLE_API = "1.8";
export const TABLES_NEED_1_8 = "Tables need PowerPoint 2021 or Microsoft 365.";
// A floor, not a promise: PowerPoint grows a row to fit what is in it.
const ROW_HEIGHT = 18;
// How many cells one round trip formats. PowerPoint for the web spent about
// 0.4 s per cell property write on 09.09, and a repaint writes five per cell,
// so eight cells keep a sync near 15 s, well inside SYNC_TIMEOUT_MS, where a
// whole 6x4 table in one batch ran past the deadline on the rig.
export const CELLS_PER_SYNC = 8;
// Excel's own default column width, in points: what a hidden column (which
// Excel reports as zero) takes beside visible ones, so it neither vanishes
// nor squeezes the others.
const HIDDEN_COLUMN_WIDTH = 48;
const ALIGNMENT = { l: "Left", c: "Center", r: "Right" } as const;

export function requireTableApi(): void {
  if (!hasPowerPointApi(TABLE_API)) throw new Error(TABLES_NEED_1_8);
}

// The cap Excel refuses an export at, checked again on the way in: a payload
// from another build - or a relay row nobody in this deck exported - would
// otherwise be written a round trip per eight cells, 161 of them for a
// 61 x 21 grid. The sentence is the one Excel already shows.
export function requireTableSize(payload: TablePayload): void {
  if (overTableCap(payload.rows, payload.cols)) throw new Error(TABLE_TOO_BIG);
}

// As wide as the source columns are, capped to the slide's content width.
// Excel reports zero for a hidden column, so a source whose columns are all
// hidden adds up to nothing: that takes the content width instead, because a
// table zero points wide cannot be seen or picked up again. columnWidths
// divides it evenly, the same way it already refuses to divide by zero.
// The widths a table is built from: a source with every column hidden stays
// as it is (the content width in even columns), one with a hidden column
// beside visible ones gives that column Excel's default width.
function sourceWidths(payload: TablePayload): number[] {
  if (!payload.widths.some((one) => one > 0)) return payload.widths;
  return payload.widths.map((one) => (one > 0 ? one : HIDDEN_COLUMN_WIDTH));
}

export function tableSize(payload: TablePayload): Size {
  const width = sourceWidths(payload).reduce((total, one) => total + one, 0);
  return {
    width: width > 0 ? Math.min(width, CONTENT_WIDTH) : CONTENT_WIDTH,
    height: payload.rows * ROW_HEIGHT,
  };
}

export async function insertTable(
  stage: string,
  item: InboxItem,
  payload: TablePayload,
  tag: LinkTag,
  target: InsertTarget = DEFAULT_TARGET,
): Promise<InsertResult> {
  requireTableApi();
  requireTableSize(payload);
  const placed = await PowerPoint.run(async (context) => {
    const resolved = await resolveTarget(
      context,
      stage,
      target,
      tableSize(payload),
    );
    const { slideId, placement, consume } = resolved;
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const shape = addTable(shapes, payload, placement.box, item.label);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    // The table's id is only known once this sync answers; a host that
    // swallows it never confirms one, so there is nothing here for a
    // cleanup to delete.
    await withSyncDeadline(context.sync(), "inserting the table");
    // The formats come once the table exists, a few cells per round trip; a
    // round trip that fails or stops answering takes the half-formatted
    // table down again, so nothing is left behind and the item stays waiting.
    try {
      const table = shape.getTable();
      await writeCellsInChunks(context, table, payload, false, FORMATTING);
    } catch (error) {
      await cleanupShapes(slideId, [shape.id]);
      throw error;
    }
    return {
      slideId,
      shapeId: shape.id,
      overlapping: placement.overlapping,
      consume,
    };
  });
  // Only reached once the table exists and is fully formatted: a consumed
  // placeholder is never dropped for an insert that ended up failing.
  await finishTarget(target, placed.slideId, placed.consume);
  return {
    slideId: placed.slideId,
    shapeId: placed.shapeId,
    overlapping: placed.overlapping,
  };
}

// The cells the source has now, written where the table already sits. Nothing
// here moves the shape: the size only decides whether it can be rewritten at
// all, and a table whose grid changed is built again at the same corner.
export async function refreshTable(
  found: FoundLink,
  payload: TablePayload,
  tag: LinkTag,
): Promise<void> {
  requireTableApi();
  requireTableSize(payload);
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  await PowerPoint.run(async (context) => {
    const shape = shapeAt(context, found);
    const table = shape.getTable();
    table.load("rowCount,columnCount");
    await withSyncDeadline(context.sync(), "reading the table");
    if (table.rowCount === payload.rows && table.columnCount === payload.cols) {
      // The tag travels with the last chunk: a repaint the host stops midway
      // keeps the old revision, so the row stays "Update available" and the
      // next update writes every cell again.
      await writeCellsInChunks(context, table, payload, true, REPAINTING, () =>
        shape.tags.add(TAG_LINK, encodeTag(tag)),
      );
      return;
    }
    const built = recreate(context, shape, found, payload, stage);
    await withSyncDeadline(context.sync(), "rebuilding the table");
    // The same rule as the repaint above: the new revision is written only
    // once every cell of it is, so a format round trip the host swallows
    // leaves a row that still says "Update available" rather than a half
    // formatted table the pane calls current and no later press finishes.
    await writeCellsInChunks(
      context,
      built.getTable(),
      payload,
      false,
      FORMATTING,
      () => {
        built.tags.add(TAG_LINK, encodeTag(tag));
      },
    );
  });
}

// A table cannot grow a row through the API, so it is built again at the
// corner and width the user left it at and tagged, and the old one is deleted
// in the same batch. The add is queued first, so a host that refuses it never
// gets as far as taking the old table - and both its tags - down. The tags it
// carries out of here are the ones the deck already held: the new revision is
// the caller's to write once the formats have landed.
function recreate(
  context: PowerPoint.RequestContext,
  old: PowerPoint.Shape,
  found: FoundLink,
  payload: TablePayload,
  stage: string,
): PowerPoint.Shape {
  if (isGrouped(found)) {
    throw new Error(`${stage}: ungroup the table before its size can change`);
  }
  const box = {
    left: found.left,
    top: found.top,
    width: found.width,
    height: tableSize(payload).height,
  };
  const label = sourceLabel(found.tag.src, found.tag.kind);
  const shapes = context.presentation.slides.getItem(found.slideId).shapes;
  const shape = addTable(shapes, payload, box, label);
  shape.tags.add(TAG_LINK, encodeTag(found.tag));
  shape.tags.add(TAG_KEY, found.token);
  old.delete();
  return shape;
}

function addTable(
  shapes: PowerPoint.ShapeCollection,
  payload: TablePayload,
  box: Box,
  label: string,
): PowerPoint.Shape {
  const shape = shapes.addTable(payload.rows, payload.cols, {
    ...box,
    values: payload.cells.map((row) => row.map((cell) => cell.t)),
    columns: columnWidths(payload, box.width).map((width) => ({
      columnWidth: width,
    })),
  });
  shape.name = `pls,fix table ${label}`;
  return shape;
}

// The source columns' widths, scaled together so they add up to the table's
// width: PowerPoint would otherwise divide the width evenly and wrap the
// label column into five lines.
export function columnWidths(payload: TablePayload, width: number): number[] {
  const source = sourceWidths(payload);
  const total = source.reduce((sum, one) => sum + one, 0);
  if (total <= 0) return payload.widths.map(() => width / payload.cols);
  // Whole points only: PowerPoint for the web refuses a fractional width.
  // The widest column absorbs the rounding remainder, so the columns still
  // add up to the table and no narrow one is squeezed to nothing.
  const scale = width / total;
  const widths = source.map((one) => Math.max(1, Math.round(one * scale)));
  const remainder =
    Math.round(width) - widths.reduce((sum, one) => sum + one, 0);
  const widest = widths.indexOf(Math.max(...widths));
  widths[widest] = Math.max(1, (widths[widest] ?? 1) + remainder);
  return widths;
}

const FORMATTING = "formatting the table";
const REPAINTING = "repainting the table";

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

// The cells one repaint or format pass writes, CELLS_PER_SYNC per round trip,
// so no batch is big enough to reach the deadline and one the host swallows
// loses a chunk rather than the whole table. `beforeLast` is queued into the
// final round trip, for a tag that may not land before every cell has.
async function writeCellsInChunks(
  context: PowerPoint.RequestContext,
  table: PowerPoint.Table,
  payload: TablePayload,
  withText: boolean,
  what: string,
  beforeLast?: () => void,
): Promise<void> {
  const chunks = cellChunks(payload, withText);
  if (chunks.length === 0) {
    if (!beforeLast) return;
    beforeLast();
    await withSyncDeadline(context.sync(), what);
    return;
  }
  for (const [index, chunk] of chunks.entries()) {
    for (const [row, column, cell] of chunk) {
      writeCell(table.getCellOrNullObject(row, column), cell, withText);
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
): void {
  if (withText) target.text = cell.t;
  // A repaint has to undo what the last one wrote; an insert only ever adds.
  if (withText || cell.b !== undefined) target.font.bold = cell.b === true;
  if (withText || cell.i !== undefined) target.font.italic = cell.i === true;
  if (withText || cell.a !== undefined) {
    target.horizontalAlignment = ALIGNMENT[cell.a ?? "l"];
  }
  if (cell.f !== undefined) target.fill.setSolidColor(cell.f);
  else if (withText) target.fill.clear();
  // Colour and size have no "back to the table style" to write, so a cell that
  // stops naming them keeps what it had until the table is built again.
  if (cell.c !== undefined) target.font.color = cell.c;
  if (cell.z !== undefined) target.font.size = cell.z;
}
