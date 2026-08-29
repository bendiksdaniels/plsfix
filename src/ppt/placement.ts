// Which slide an inserted link goes on and where it lands: PowerPoint's own
// selection, the boxes the slide's other shapes occupy, and the first free
// spot placeInFreeSpace finds. Owns the slide geometry constants and the
// empty-placeholder rule. Invariant: an empty layout placeholder never counts
// as occupied and every other shape does, so a second insert cannot land on
// the first.

import {
  placeInFreeSpace,
  type Box,
  type Placement,
  type Size,
} from "../layout";
import { SLIDE_16_9 } from "../link/status";
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
  await context.sync();
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

// The box for a new object of this size on that slide. Two syncs: the slide's
// shapes, then - only when the slide has placeholders - whether each holds
// text. Never writes anything; the caller creates the shape at the box.
export async function placeOnSlide(
  context: PowerPoint.RequestContext,
  slideId: string,
  size: Size,
): Promise<Placement> {
  const occupied = await occupiedBoxes(context, slideId);
  return placeInFreeSpace(size, occupied, SLIDE, SLIDE_MARGIN, SLIDE_GAP);
}

async function occupiedBoxes(
  context: PowerPoint.RequestContext,
  slideId: string,
): Promise<Box[]> {
  const shapes = context.presentation.slides.getItem(slideId).shapes;
  shapes.load(SHAPE_PROPERTIES);
  await context.sync();
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
  await context.sync();
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
