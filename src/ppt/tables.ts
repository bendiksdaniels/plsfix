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
  TAG_PAINT,
  type InboxItem,
  type LinkTag,
  type TablePayload,
} from "../link/model";
import { encodePaintMap } from "../link/paint-map";
import { cleanupByToken, cleanupShapes } from "./chart-cleanup";
import { withSyncDeadline } from "./chart-draw";
import type { FoundLink, InsertResult } from "./host";
import {
  CONTENT_WIDTH,
  DEFAULT_TARGET,
  finishTarget,
  resolveTarget,
  type InsertTarget,
} from "./placement";
import { awaitShapeRegistered } from "./shape-ready";
import { hasPowerPointApi, isGrouped, shapeAt } from "./shapes";
import {
  CELLS_PER_SYNC,
  FORMATTING,
  NO_CLEARS,
  REPAINTING,
  writeCellsInChunks,
} from "./table-cells";
import { uniformSize } from "./table-font";
import { fillsToClear, queueTableStyle, readPaintTag } from "./table-style";

export { CELLS_PER_SYNC };

// shapes.addTable and the Table object both arrived in PowerPointApi 1.8.
const TABLE_API = "1.8";
export const TABLES_NEED_1_8 = "Tables need PowerPoint 2021 or Microsoft 365.";
// A floor, not a promise: PowerPoint grows a row to fit what is in it.
const ROW_HEIGHT = 18;
// Excel's own default column width, in points: what a hidden column (which
// Excel reports as zero) takes beside visible ones, so it neither vanishes
// nor squeezes the others.
const HIDDEN_COLUMN_WIDTH = 48;

export function requireTableApi(): void {
  if (!hasPowerPointApi(TABLE_API)) throw new Error(TABLES_NEED_1_8);
}

// One wording per stage, for its withSyncDeadline and awaitShapeRegistered.
const INSERTING_TABLE = "inserting the table";
const REBUILDING_TABLE = "rebuilding the table";

// The cap Excel refuses an export at, checked again on the way in: a payload
// from another build - or a relay row nobody in this deck exported - would
// otherwise be written a round trip per eight cells, 161 of them for a
// 61 x 21 grid. The sentence is the one Excel already shows.
//
// overTableCap (link/model.ts) only ever checks the ceiling: a hand-crafted
// or corrupted payload with 0 rows or 0 columns (or a negative/non-integer
// count JSON can carry straight through) sails past it and would reach
// shapes.addTable, which throws a raw InvalidArgument PowerPoint's own way -
// exactly what this guard exists to keep out of the toast. Checked first, so
// the too-big and too-small refusals never share one sentence.
const TABLE_TOO_SMALL = "Tables need at least one row and one column.";

function invalidTableExtent(count: number): boolean {
  return !Number.isInteger(count) || count < 1;
}

export function requireTableSize(payload: TablePayload): void {
  if (invalidTableExtent(payload.rows) || invalidTableExtent(payload.cols)) {
    throw new Error(TABLE_TOO_SMALL);
  }
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
    const { slideId, placement, consume, before } = resolved;
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const shape = addTable(shapes, payload, placement.box, item.label);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    // The table's id is only known once this sync answers, so a host that
    // swallows it (rolls the batch back) leaves nothing to delete - but one
    // that applies the add and both tags before answering the sync itself
    // with an error leaves a fully tagged, orphaned table this catch is
    // what finds and removes. cleanupByToken (chart-cleanup.ts) matches by
    // the token, which the link owns, not this one push - `before` keeps
    // it off an older, already-finished copy of the link.
    try {
      await withSyncDeadline(context.sync(), INSERTING_TABLE);
    } catch (error) {
      await withSyncDeadline(cleanupByToken(slideId, item.token, before)).catch(
        () => undefined,
      );
      throw error;
    }
    await formatNewTable(context, shapes, slideId, shape, payload);
    return {
      slideId,
      shapeId: shape.id,
      overlapping: placement.overlapping,
      freeSpot: placement.freeSpot,
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
    freeSpot: placed.freeSpot,
  };
}

// The formats come once the table exists, a few cells per round trip; a
// round trip that fails or stops answering takes the half-formatted table
// down again, so nothing is left behind and the item stays waiting. It first
// waits out the web's registration window (shape-ready.ts, rig 27.09):
// getTable() goes through getItem(id), which fails 5010 right after the add.
async function formatNewTable(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  slideId: string,
  shape: PowerPoint.Shape,
  payload: TablePayload,
): Promise<void> {
  try {
    await awaitShapeRegistered(context, shapes, shape.id, INSERTING_TABLE);
    const table = shape.getTable();
    queueTableStyle(table, payload.h === true);
    await writeCellsInChunks(context, table, payload, {
      withText: false,
      clear: NO_CLEARS,
      what: FORMATTING,
      uniformFontSize: uniformSize(payload),
      beforeLast: () =>
        shape.tags.add(TAG_PAINT, encodePaintMap(payload.cells)),
    });
  } catch (error) {
    await withSyncDeadline(cleanupShapes(slideId, [shape.id])).catch(
      () => undefined,
    );
    throw error;
  }
}

// What a shape's tags cost to read: every key and value in one round trip,
// the same load host.ts's scan uses, so TAG_PAINT rides the size check below
// rather than spending a sync of its own.
const TAG_PROPERTIES = "items/key,items/value";

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
    shape.tags.load(TAG_PROPERTIES);
    await withSyncDeadline(context.sync(), "reading the table");
    if (table.rowCount === payload.rows && table.columnCount === payload.cols) {
      await repaintInPlace(context, shape, table, payload, tag);
      return;
    }
    await rebuildTable(context, shape, found, payload, tag, stage);
  });
}

// The header flag is queued blind, with no read first: a repaint costs one
// round trip whether or not the source carries a header row. The tag and the
// paint map travel with the last chunk, same as the link tag always has, so a
// repaint the host stops midway keeps the old revision and the next update
// writes - and clears - every cell again.
async function repaintInPlace(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  table: PowerPoint.Table,
  payload: TablePayload,
  tag: LinkTag,
): Promise<void> {
  queueTableStyle(table, payload.h === true);
  const clear = fillsToClear(readPaintTag(shape), payload);
  await writeCellsInChunks(context, table, payload, {
    withText: true,
    clear,
    what: REPAINTING,
    beforeLast: () => {
      shape.tags.add(TAG_LINK, encodeTag(tag));
      shape.tags.add(TAG_PAINT, encodePaintMap(payload.cells));
    },
  });
}

// The same rule as the repaint above: the new revision is written only once
// every cell of it is, so a format round trip the host swallows leaves a row
// that still says "Update available" rather than a half formatted table the
// pane calls current and no later press finishes. Same registration window
// as formatNewTable's, one sync later: built.getTable() re-anchors the same
// way once recreate's own add-sync has loaded its id (rig 27.09).
async function rebuildTable(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  found: FoundLink,
  payload: TablePayload,
  tag: LinkTag,
  stage: string,
): Promise<void> {
  const built = recreate(context, shape, found, payload, stage);
  await withSyncDeadline(context.sync(), REBUILDING_TABLE);
  const shapes = context.presentation.slides.getItem(found.slideId).shapes;
  await awaitShapeRegistered(context, shapes, built.id, REBUILDING_TABLE);
  const table = built.getTable();
  queueTableStyle(table, payload.h === true);
  await writeCellsInChunks(context, table, payload, {
    withText: false,
    clear: NO_CLEARS,
    what: FORMATTING,
    uniformFontSize: uniformSize(payload),
    beforeLast: () => {
      built.tags.add(TAG_LINK, encodeTag(tag));
      built.tags.add(TAG_PAINT, encodePaintMap(payload.cells));
    },
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
  // rebuildTable polls shapes.getItemOrNullObject(built.id) once this loads.
  shape.load("id");
  old.delete();
  return shape;
}

// PowerPoint sizes every row for the table's DEFAULT font (18 pt) the moment
// values land, and never shrinks a row once a smaller size arrives afterwards
// - which is why a freshly inserted table came out double height until the
// first repaint. Asking for the payload's own modal size up front means the
// table is never sized for a font it will not keep.
function addTable(
  shapes: PowerPoint.ShapeCollection,
  payload: TablePayload,
  box: Box,
  label: string,
): PowerPoint.Shape {
  const size = uniformSize(payload);
  // Rounded once, here, and nowhere else: the table's own width and its
  // columns' widths both have to come from this one whole-point number, or
  // PowerPoint for the web refuses the add (see columnWidths below). box.width
  // arrives fractional whenever fitInto placed this table (a Spot or the
  // selected shape, or recreate() below reading an existing shape's own
  // possibly-fractional width) - free-space placement already rounds its own
  // box, but this rounding is unconditional, so both boxes end up exact.
  const width = Math.round(box.width);
  const shape = shapes.addTable(payload.rows, payload.cols, {
    ...box,
    width,
    values: payload.cells.map((row) => row.map((cell) => cell.t)),
    columns: columnWidths(payload, width).map((one) => ({
      columnWidth: one,
    })),
    ...(size === undefined
      ? {}
      : { uniformCellProperties: { font: { size } } }),
  });
  shape.name = `pls,fix ${label}`;
  return shape;
}

// The source columns' widths, scaled together so they add up to the table's
// width: PowerPoint would otherwise divide the width evenly and wrap the
// label column into five lines. A source with nothing to scale by (every
// column hidden, so sourceWidths hands back the all-zero widths unchanged)
// shares the width evenly instead - equal weights through the same one path,
// rather than a second formula that could drift from the first.
export function columnWidths(payload: TablePayload, width: number): number[] {
  const source = sourceWidths(payload);
  const total = source.reduce((sum, one) => sum + one, 0);
  const weights = total > 0 ? source : source.map(() => 1);
  return scaleToWhole(weights, width);
}

// Positive weights scaled to whole points that add up to width, rounded to
// the nearest point ONCE, here. PowerPoint for the web refuses
// shapes.addTable with InvalidArgument unless the columns' columnWidth
// values add up EXACTLY to the table's own width (probed 27.09); fractional
// columns summing exactly to a fractional width are accepted, so that
// mismatch is the cause, not fractions as such - which is why addTable above
// rounds the table's own width through this same Math.round before it ever
// reaches here, rather than leaving one side of the comparison fractional.
// Every column keeps a floor of one point first (Excel's narrowest column
// still gets a place in the table), and what is left of the width past that
// floor is apportioned by weight, largest-remainder style: floor each share,
// then hand the units the floors still owe to whichever columns lost the most
// to their own floor, ties by position. That spreads a rounding bias across
// as many columns as it takes, rather than one column absorbing all of it and
// hitting its own floor before the sum is exact (property-tested, hunt2).
function scaleToWhole(weights: number[], width: number): number[] {
  const cols = weights.length;
  if (cols === 0) return [];
  const rounded = Math.round(width);
  const total = weights.reduce((sum, one) => sum + one, 0);
  const shares =
    total > 0 ? weights.map((one) => one / total) : weights.map(() => 1 / cols);
  // Never negative: a table narrower than one point per column has nothing
  // left to apportion once every column has already spent its own floor.
  const budget = Math.max(0, rounded - cols);
  const ideal = shares.map((share) => share * budget);
  const floor = ideal.map((one) => Math.floor(one));
  const flooredSum = floor.reduce((sum, one) => sum + one, 0);
  const owed = Math.max(0, Math.min(cols, budget - flooredSum));
  const bumpOrder = ideal
    .map((one, index) => ({ remainder: one - floor[index]!, index }))
    .sort((a, b) => b.remainder - a.remainder || a.index - b.index)
    .map((entry) => entry.index);
  const bumped = new Set(bumpOrder.slice(0, owed));
  return floor.map((one, index) => 1 + one + (bumped.has(index) ? 1 : 0));
}
