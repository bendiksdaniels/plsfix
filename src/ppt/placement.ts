// Which slide an inserted link goes on and where it lands: PowerPoint's own
// selection, the boxes other shapes occupy and the first free spot
// placeInFreeSpace finds, a named spot (spotBox + fitInto), or the first
// selected shape's box. Invariant: an empty placeholder and a dashed empty
// frame never count as occupied; a group with no box of its own occupies
// the union of its children; resolveTarget's "select a slide/shape first"
// is the only raw error surfaced. It also hands back `before`, the slide's
// pre-add ids, which cleanupByToken trusts to spare a shape already there.

import { placeInFreeSpace } from "../free-space";
import {
  fitInto,
  hasArea,
  isDecorativeFrame,
  reportedBox,
  spotBox,
  type Box,
  type Placement,
  type Size,
  type Spot,
} from "../layout";
import { SLIDE_16_9 } from "../link/status";
import { textOccupiedHeight } from "../text-extent";
import { withSyncDeadline } from "./chart-draw";
import { requireObjectToolsApi } from "./object-tools";
import {
  GROUP_API,
  GROUP_TYPE,
  hasPowerPointApi,
  SHAPE_PROPERTIES,
} from "./shapes";

// Half an inch of margin, the same one fitToSlide keeps, and a gap wide enough
// that two objects beside each other read as two.
export const SLIDE_MARGIN = 36;
export const SLIDE_GAP = 12;
// PowerPoint reports a layout placeholder as its own shape type; an empty one
// is the slide's "click to add" furniture, not an object to place around.
const PLACEHOLDER = "Placeholder";
// A plain text box: the other kind whose occupied box is cut to its text.
// textFrame throws InvalidArgument on a shape without one (picture, table,
// group), so the read below only ever targets these two types.
const TEXT_BOX = "TextBox";
const TEXT_SHAPE_TYPES: ReadonlySet<string> = new Set([PLACEHOLDER, TEXT_BOX]);

export const SLIDE = SLIDE_16_9;
export const CONTENT_WIDTH = SLIDE.width - 2 * SLIDE_MARGIN;

// The slide PowerPoint reports as selected; null when there is none, so the
// caller can say so instead of guessing one.
export async function readSelectedSlideId(
  context: PowerPoint.RequestContext,
): Promise<string | null> {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await withSyncDeadline(context.sync(), "reading the slide");
  return selected.items[0]?.id ?? null;
}

async function selectedSlideId(
  context: PowerPoint.RequestContext,
  stage: string,
): Promise<string> {
  const id = await readSelectedSlideId(context);
  if (!id) throw new Error(`${stage}: select a slide first.`);
  return id;
}

// "free" keeps the old placeInFreeSpace search; a Spot is one of layout.ts's
// named halves/quarters/whole; "selected-shape" fits the insert into
// PowerPoint's own current selection instead of hunting for room.
export type Where = "free" | "selected-shape" | Spot;

// What the inbox's Slide and Where pickers resolve to. slideId null means
// "This slide", the picker's default - resolveTarget reads the active one the
// same way every insert always has.
export interface InsertTarget {
  slideId: string | null;
  where: Where;
}

// Every existing insert call keeps behaving exactly as it did before the
// pickers existed.
export const DEFAULT_TARGET: InsertTarget = { slideId: null, where: "free" };

export interface ResolvedTarget {
  slideId: string;
  placement: Placement;
  // Set only for "selected-shape" over an empty layout placeholder: the
  // caller deletes it once the new shape is confirmed, so the placeholder is
  // replaced rather than left sitting under the new object.
  consume?: string;
  // The target slide's own top-level ids, read before this insert added
  // anything: cleanupByToken (chart-cleanup.ts) cleans up by the link's
  // token, not this push's, so a pre-existing copy is never mistaken for it.
  before: ReadonlySet<string>;
}

const NO_SHAPE_SELECTED =
  "Select a shape on the slide first, or choose another spot.";

// The one place every insert decides its slide and box: target.slideId when
// the picker named one, the selected slide otherwise; "free" is the old
// placeInFreeSpace search, a Spot fits size into spotBox (never flagged as
// overlapping - the user chose that box), "selected-shape" reads PowerPoint's
// current selection. minScale only ever reaches "free".
export async function resolveTarget(
  context: PowerPoint.RequestContext,
  stage: string,
  target: InsertTarget,
  size: Size,
  minScale?: number,
): Promise<ResolvedTarget> {
  const slideId = target.slideId ?? (await selectedSlideId(context, stage));
  if (target.where === "free") {
    const { placement, before } = await placeOnSlide(
      context,
      slideId,
      size,
      minScale,
    );
    return { slideId, placement, before };
  }
  if (target.where === "selected-shape") {
    return { slideId, ...(await selectedShapeTarget(context, slideId, size)) };
  }
  const spot = await spotTarget(context, slideId, target.where, size);
  return { slideId, ...spot };
}

// Runs once an insert's shape id is confirmed: drops the placeholder
// resolveTarget flagged for consumption, and - only when the picker named
// an explicit slide - brings that slide on screen ("This slide" leaves the
// view exactly where it was). One or two more round trips, none when
// neither has anything to do. getItemOrNullObject makes a consume id gone
// by now (already deleted, or a slide-scoped id that never named a shape
// here - the bug this guards against) a no-op, never a raw ItemNotFound
// after the insert has already landed.
export async function finishTarget(
  target: InsertTarget,
  slideId: string,
  consume?: string,
): Promise<void> {
  if (consume === undefined && target.slideId === null) return;
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const maybeConsumed =
      consume === undefined ? null : shapes.getItemOrNullObject(consume);
    if (maybeConsumed) {
      maybeConsumed.load("isNullObject");
      await withSyncDeadline(context.sync(), "finding the placeholder");
      if (!maybeConsumed.isNullObject) maybeConsumed.delete();
    }
    if (target.slideId !== null)
      context.presentation.setSelectedSlides([slideId]);
    await withSyncDeadline(context.sync(), "finishing the insert");
  });
}

// The box for a new object of this size on that slide. Two syncs: the slide's
// shapes, then - only when the slide has placeholders - whether each holds
// text. Never writes anything; the caller creates the shape at the box.
async function placeOnSlide(
  context: PowerPoint.RequestContext,
  slideId: string,
  size: Size,
  // How far the free-space scan may shrink the object before it gives up and
  // centres it over what is there: a chart passes the scale that would take it
  // to MIN_SIZE, because a group of hairlines is worse than an overlap.
  minScale?: number,
): Promise<{ placement: Placement; before: ReadonlySet<string> }> {
  const { boxes: occupied, before } = await occupiedBoxes(context, slideId);
  const placement = placeInFreeSpace(
    size,
    occupied,
    SLIDE,
    SLIDE_MARGIN,
    SLIDE_GAP,
    minScale,
  );
  return { placement, before };
}

async function occupiedBoxes(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<{ boxes: Box[]; before: ReadonlySet<string> }> {
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  shapes.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading the slide's shapes");
  // ResolvedTarget.before: every id already on the slide.
  const before = new Set(shapes.items.map((shape) => shape.id));
  const empty = await emptyPlaceholders(context, shapes.items);
  await loadFrameLooks(context, shapes.items);
  const extents = await loadTextExtents(context, shapes.items);
  const boxes: Box[] = [];
  for (const shape of shapes.items) {
    if (empty.has(shape.id) || isFrame(shape)) continue;
    const box = await boundsOf(context, shape);
    if (!hasArea(box)) continue;
    const extent = extents.get(shape.id);
    boxes.push(extent ? trimToText(box, extent) : box);
  }
  return { boxes, before };
}

interface TextExtent {
  text: string;
  fontSize: number;
}

// A placeholder or text box's own text and font size, for the shapes whose
// text is worth reading at all: not empty (an empty one is already excluded
// above, or - a plain text box - stays fully occupied, today's behaviour) and
// of a type that supports a text frame in the first place.
async function loadTextExtents(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<Map<string, TextExtent>> {
  const candidates = shapes.filter((shape) => TEXT_SHAPE_TYPES.has(shape.type));
  if (candidates.length === 0) return new Map();
  for (const shape of candidates) {
    shape.textFrame.textRange.load("text");
    shape.textFrame.textRange.font.load("size");
  }
  await withSyncDeadline(context.sync(), "reading the text");
  const extents = new Map<string, TextExtent>();
  for (const shape of candidates) {
    const text = shape.textFrame.textRange.text;
    const fontSize = shape.textFrame.textRange.font.size;
    // Empty text is excluded here (an empty placeholder is already free, and
    // a plain empty text box stays fully occupied, today's behaviour); a
    // size that is not a plain number - text set in more than one size, or a
    // host that never reports one - is left untrimmed rather than guessed
    // at, the same conservative answer as no extent at all.
    if (text.length === 0 || typeof fontSize !== "number") continue;
    extents.set(shape.id, { text, fontSize });
  }
  return extents;
}

// The box's own left/top/width unchanged, height cut to what the text needs:
// a placeholder occupies the lines it holds, not the whole frame it was
// given, so free space below a short caption is free space again.
function trimToText(box: Box, extent: TextExtent): Box {
  const height = textOccupiedHeight(
    extent.text,
    extent.fontSize,
    box.width,
    box.height,
  );
  return { ...box, height };
}

const FRAME_TYPE = "GeometricShape";

async function loadFrameLooks(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<void> {
  const frames = shapes.filter((shape) => shape.type === FRAME_TYPE);
  if (frames.length === 0) return;
  for (const shape of frames) {
    shape.fill.load("type");
    shape.lineFormat.load("visible,dashStyle");
    shape.textFrame.load("hasText");
  }
  await withSyncDeadline(context.sync(), "reading the outlines");
}

function isFrame(shape: PowerPoint.Shape): boolean {
  if (shape.type !== FRAME_TYPE) return false;
  return isDecorativeFrame({
    fillType: String(shape.fill.type),
    hasText: Boolean(shape.textFrame.hasText),
    dashStyle:
      shape.lineFormat.dashStyle == null
        ? null
        : String(shape.lineFormat.dashStyle),
    lineVisible: Boolean(shape.lineFormat.visible),
  });
}

async function boundsOf(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
): Promise<Box> {
  const own = boxOf(shape);
  if (
    hasArea(own) ||
    shape.type !== GROUP_TYPE ||
    !hasPowerPointApi(GROUP_API)
  ) {
    return own;
  }
  const inner = shape.group.shapes;
  inner.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading a group's shapes");
  const children: Box[] = [];
  for (const child of inner.items) {
    children.push(await boundsOf(context, child));
  }
  return reportedBox(own, children);
}

// hasText is asked of placeholders alone: reading a text frame that does
// not apply (a picture, a table) is how a batch fails on a real host.
async function emptyPlaceholders(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<Set<string>> {
  const placeholders = shapes.filter((shape) => shape.type === PLACEHOLDER);
  if (placeholders.length === 0) return new Set();
  for (const shape of placeholders) shape.textFrame.load("hasText");
  await withSyncDeadline(context.sync(), "reading the placeholders");
  return new Set(
    placeholders
      .filter((shape) => !shape.textFrame.hasText)
      .map((shape) => shape.id),
  );
}

function boxOf(shape: PowerPoint.Shape): Box {
  return {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

// fitInto never enlarges: the box it hands back is size scaled by at most
// 1, the same meaning placeInFreeSpace gives Placement.scale.
function scaleOf(size: Size, box: Box): number {
  return size.width > 0 ? box.width / size.width : 1;
}

// A named spot: fitInto keeps the box's shape, and landing on something
// there is not the free-space failure OVERLAP_NOTE describes - overlapping
// stays false. The one new round trip is cleanupByToken's `before` read.
async function spotTarget(
  context: PowerPoint.RequestContext,
  slideId: string,
  spot: Spot,
  size: Size,
): Promise<{ placement: Placement; before: ReadonlySet<string> }> {
  const box = fitInto(size, spotBox(spot, SLIDE, SLIDE_MARGIN, SLIDE_GAP));
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  shapes.load("items/id");
  await withSyncDeadline(context.sync(), "reading the slide's shapes");
  return {
    placement: { box, scale: scaleOf(size, box), overlapping: false },
    before: new Set(shapes.items.map((shape) => shape.id)),
  };
}

// PowerPoint's current selection, gated like the Tools tab's own object
// tools: the first selected shape is the anchor, but only when it sits on
// the target slide - PowerPoint shape ids are unique per slide, so a
// selection left over from another slide must be refused here rather than
// trusted into this slide's geometry (and handed to finishTarget as a
// consume id that could name an unrelated shape). An empty layout
// placeholder is consumed so the new object replaces it; anything else
// stays, deliberately overlapped - the user's own choice, not the "no free
// space" OVERLAP_NOTE describes.
async function selectedShapeTarget(
  context: PowerPoint.RequestContext,
  slideId: string,
  size: Size,
): Promise<{
  placement: Placement;
  consume?: string;
  before: ReadonlySet<string>;
}> {
  requireObjectToolsApi();
  const selected = context.presentation.getSelectedShapes();
  selected.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading the selection");
  const shape = selected.items[0];
  if (!shape) throw new Error(NO_SHAPE_SELECTED);
  // Three things queued into this one sync: the parent slide's id (is the
  // selection on the target slide), a placeholder's own text flag, and the
  // slide's top-level ids for cleanupByToken's `before`.
  const parentSlide = shape.getParentSlide();
  parentSlide.load("id");
  const isPlaceholder = shape.type === PLACEHOLDER;
  if (isPlaceholder) shape.textFrame.load("hasText");
  const slideShapes = context.presentation.slides.getItem(slideId).shapes;
  slideShapes.load("items/id");
  await withSyncDeadline(context.sync(), "reading the selected shape's slide");
  if (parentSlide.id !== slideId) throw new Error(NO_SHAPE_SELECTED);
  const consume =
    isPlaceholder && !shape.textFrame.hasText ? shape.id : undefined;
  const box = fitInto(size, boxOf(shape));
  return {
    placement: { box, scale: scaleOf(size, box), overlapping: false },
    consume,
    before: new Set(slideShapes.items.map((one) => one.id)),
  };
}
