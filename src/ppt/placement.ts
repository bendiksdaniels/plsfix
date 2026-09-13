// Which slide an inserted link goes on and where it lands: PowerPoint's own
// selection, the boxes the slide's other shapes occupy, and the first free
// spot placeInFreeSpace finds - or, since the inbox grew a Slide/Where picker,
// a named spot (spotBox + fitInto) or the first selected shape's box. Owns the
// slide geometry constants and the empty-placeholder rule. Invariant: an
// empty layout placeholder never counts as occupied and every other shape
// does, so a second insert cannot land on the first, and resolveTarget's
// "select a slide/shape first" errors are the only ones its callers see for
// those cases - never a raw office.js string.

import {
  fitInto,
  overlaps,
  placeInFreeSpace,
  spotBox,
  type Box,
  type Placement,
  type Size,
  type Spot,
} from "../layout";
import { SLIDE_16_9 } from "../link/status";
import { withSyncDeadline } from "./chart-draw";
import { requireObjectToolsApi } from "./object-tools";
import { SHAPE_PROPERTIES } from "./shapes";

// Half an inch of margin, the same one fitToSlide keeps, and a gap wide enough
// that two objects beside each other read as two.
const SLIDE_MARGIN = 36;
const SLIDE_GAP = 12;
// PowerPoint reports a layout placeholder as its own shape type; an empty one
// is the slide's "click to add" furniture, not an object to place around.
const PLACEHOLDER = "Placeholder";

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

export async function selectedSlideId(
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
}

const NO_SHAPE_SELECTED =
  "Select a shape on the slide first, or choose another spot.";

// The one place every insert decides its slide and box: target.slideId when
// the picker named one, the selected slide otherwise (selectedSlideId's own
// "select a slide first" error, unchanged); then "free" is the old
// placeInFreeSpace search, a Spot fits the size into spotBox and flags
// overlapping against what the slide already holds, and "selected-shape"
// reads PowerPoint's current selection. minScale only ever reaches "free":
// a chart passes the scale it would rather shrink to than overlap.
export async function resolveTarget(
  context: PowerPoint.RequestContext,
  stage: string,
  target: InsertTarget,
  size: Size,
  minScale?: number,
): Promise<ResolvedTarget> {
  const slideId = target.slideId ?? (await selectedSlideId(context, stage));
  if (target.where === "free") {
    const placement = await placeOnSlide(context, slideId, size, minScale);
    return { slideId, placement };
  }
  if (target.where === "selected-shape") {
    return { slideId, ...(await selectedShapeTarget(context, size)) };
  }
  const spot = await spotTarget(context, slideId, target.where, size);
  return { slideId, ...spot };
}

// Runs once an insert's shape id is confirmed: drops the placeholder
// resolveTarget flagged for consumption, and - only when the picker named an
// explicit slide, "This slide" leaves the view exactly where it was - brings
// that slide on screen the way picture.ts's selection insert already does.
// One more round trip, and only when either has something to do.
export async function finishTarget(
  target: InsertTarget,
  slideId: string,
  consume?: string,
): Promise<void> {
  if (consume === undefined && target.slideId === null) return;
  await PowerPoint.run(async (context) => {
    if (consume !== undefined) {
      context.presentation.slides
        .getItem(slideId)
        .shapes.getItem(consume)
        .delete();
    }
    if (target.slideId !== null)
      context.presentation.setSelectedSlides([slideId]);
    await withSyncDeadline(context.sync(), "finishing the insert");
  });
}

// The box for a new object of this size on that slide. Two syncs: the slide's
// shapes, then - only when the slide has placeholders - whether each holds
// text. Never writes anything; the caller creates the shape at the box.
export async function placeOnSlide(
  context: PowerPoint.RequestContext,
  slideId: string,
  size: Size,
  // How far the free-space scan may shrink the object before it gives up and
  // centres it over what is there: a chart passes the scale that would take it
  // to MIN_SIZE, because a group of hairlines is worse than an overlap.
  minScale?: number,
): Promise<Placement> {
  const occupied = await occupiedBoxes(context, slideId);
  return placeInFreeSpace(
    size,
    occupied,
    SLIDE,
    SLIDE_MARGIN,
    SLIDE_GAP,
    minScale,
  );
}

async function occupiedBoxes(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<Box[]> {
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  shapes.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading the slide's shapes");
  const empty = await emptyPlaceholders(context, shapes.items);
  return shapes.items
    .filter((shape) => !empty.has(shape.id))
    .map((shape) => boxOf(shape));
}

// hasText is asked of placeholders alone: a picture or a table has no text
// frame worth reading, and on a real host reading one that does not apply is
// how a batch fails.
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

// fitInto never enlarges, so the box it hands back is always size scaled by
// at most 1; Placement.scale is that factor, the same meaning placeInFreeSpace
// gives it.
function scaleOf(size: Size, box: Box): number {
  return size.width > 0 ? box.width / size.width : 1;
}

// A named spot: fitInto keeps the box's shape, and overlapping is the truth
// about whatever the slide already holds there - a spot is a box, not a
// guarantee, so the pane still says so when it lands on something.
async function spotTarget(
  context: PowerPoint.RequestContext,
  slideId: string,
  spot: Spot,
  size: Size,
): Promise<{ placement: Placement }> {
  const box = fitInto(size, spotBox(spot, SLIDE, SLIDE_MARGIN, SLIDE_GAP));
  const occupied = await occupiedBoxes(context, slideId);
  return {
    placement: {
      box,
      scale: scaleOf(size, box),
      overlapping: occupied.some((other) => overlaps(box, other, SLIDE_GAP)),
    },
  };
}

// PowerPoint's current selection, gated the same way the Tools tab's object
// tools already are: this picker sits beside those same actions and needs no
// API they do not already require. The first selected shape is the anchor;
// an empty layout placeholder is consumed, so the new object replaces it
// instead of sitting over it - anything else stays, and the new object is
// deliberately placed over it.
async function selectedShapeTarget(
  context: PowerPoint.RequestContext,
  size: Size,
): Promise<{ placement: Placement; consume?: string }> {
  requireObjectToolsApi();
  const selected = context.presentation.getSelectedShapes();
  selected.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading the selection");
  const shape = selected.items[0];
  if (!shape) throw new Error(NO_SHAPE_SELECTED);
  const empty = await emptyPlaceholders(context, [shape]);
  const consume = empty.has(shape.id) ? shape.id : undefined;
  const box = fitInto(size, boxOf(shape));
  return {
    placement: {
      box,
      scale: scaleOf(size, box),
      // Consumed, the placeholder cannot still be under the new object;
      // kept, the new object lands on it by the user's own choice.
      overlapping: consume === undefined,
    },
    consume,
  };
}
