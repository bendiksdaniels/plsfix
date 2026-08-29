// Native PowerPoint tables (PowerPointApi 1.8): insert a table link at a free
// spot on the slide, write every cell's text and formats, and repaint one in
// place while its dimensions still match - building it again at the same
// corner when they do not. Owns every table-shaped Office.js call.
// Invariant: a repaint writes cells, never geometry.

import type { Box, Size } from "../layout";
import {
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type TableCell,
  type TablePayload,
} from "../link/model";
import type { FoundLink, InsertResult } from "./host";
import { CONTENT_WIDTH, placeOnSlide, selectedSlideId } from "./placement";
import { hasPowerPointApi, isGrouped, shapeAt } from "./shapes";

// shapes.addTable and the Table object both arrived in PowerPointApi 1.8.
const TABLE_API = "1.8";
export const TABLES_NEED_1_8 = "Tables need PowerPoint 2021 or Microsoft 365.";
// A floor, not a promise: PowerPoint grows a row to fit what is in it.
const ROW_HEIGHT = 18;
const ALIGNMENT = { l: "Left", c: "Center", r: "Right" } as const;

export function requireTableApi(): void {
  if (!hasPowerPointApi(TABLE_API)) throw new Error(TABLES_NEED_1_8);
}

// As wide as the source columns are, capped to the slide's content width.
export function tableSize(payload: TablePayload): Size {
  const width = payload.widths.reduce((total, one) => total + one, 0);
  return {
    width: Math.min(width, CONTENT_WIDTH),
    height: payload.rows * ROW_HEIGHT,
  };
}

export async function insertTable(
  stage: string,
  item: InboxItem,
  payload: TablePayload,
  tag: LinkTag,
): Promise<InsertResult> {
  requireTableApi();
  return PowerPoint.run(async (context) => {
    const slideId = await selectedSlideId(context, stage);
    const placed = await placeOnSlide(context, slideId, tableSize(payload));
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const shape = addTable(shapes, payload, placed.box, item.label);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    await context.sync();
    return { slideId, shapeId: shape.id, overlapping: placed.overlapping };
  });
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
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  await PowerPoint.run(async (context) => {
    const shape = shapeAt(context, found);
    const table = shape.getTable();
    table.load("rowCount,columnCount");
    await context.sync();
    if (table.rowCount === payload.rows && table.columnCount === payload.cols) {
      writeCells(table, payload, true);
      shape.tags.add(TAG_LINK, encodeTag(tag));
    } else {
      recreate(context, shape, found, payload, tag, stage);
    }
    await context.sync();
  });
}

// A table cannot grow a row through the API, so it is built again at the
// corner and width the user left it at and tagged, and the old one is deleted
// in the same batch. The add is queued first, so a host that refuses it never
// gets as far as taking the old table - and both its tags - down.
function recreate(
  context: PowerPoint.RequestContext,
  old: PowerPoint.Shape,
  found: FoundLink,
  payload: TablePayload,
  tag: LinkTag,
  stage: string,
): void {
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
  shape.tags.add(TAG_LINK, encodeTag(tag));
  shape.tags.add(TAG_KEY, found.token);
  old.delete();
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
  });
  shape.name = `pls,fix table ${label}`;
  writeCells(shape.getTable(), payload, false);
  return shape;
}

// Text is written on a repaint only: an insert carries it in `values`, so a
// plain cell then costs no call at all - which is what keeps a 60x20 table one
// round trip rather than twelve hundred.
function writeCells(
  table: PowerPoint.Table,
  payload: TablePayload,
  withText: boolean,
): void {
  payload.cells.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (!withText && !formatted(cell)) return;
      writeCell(
        table.getCellOrNullObject(rowIndex, columnIndex),
        cell,
        withText,
      );
    });
  });
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
