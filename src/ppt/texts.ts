// A text link on a slide: one tagged text box of its own, auto-sized to what
// Excel displayed, placed in free space like every other kind. Owns the insert
// and the in-place refresh. Invariant: a refresh writes the text and the tag,
// never the geometry, font or colour the user chose.

import type { Size } from "../layout";
import {
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type TextPayload,
} from "../link/model";
import { withSyncDeadline } from "./chart-draw";
import type { FoundLink, InsertResult } from "./host";
import { isMissingShape, missingShapeError } from "./missing-shape";
import { CONTENT_WIDTH, placeOnSlide, selectedSlideId } from "./placement";
import { shapeAt } from "./shapes";

// PowerPoint's default text-box font is 18 pt, and 0.55 em is about the width
// of a proportional font's average glyph. The auto-size corrects the guess the
// moment the box is on the slide; this only has to be close enough that the
// placement finds it a spot of roughly the right shape.
const FONT_PT = 18;
const EM_PER_CHAR = 0.55;
const PADDING = 14;
const LINE_HEIGHT = 28;
const MIN_WIDTH = 60;

export function textSize(payload: TextPayload): Size {
  const width = Math.ceil(payload.text.length * FONT_PT * EM_PER_CHAR);
  return {
    width: Math.min(CONTENT_WIDTH, Math.max(MIN_WIDTH, width + PADDING)),
    height: LINE_HEIGHT,
  };
}

export async function insertText(
  stage: string,
  item: InboxItem,
  payload: TextPayload,
  tag: LinkTag,
): Promise<InsertResult> {
  return PowerPoint.run(async (context) => {
    const slideId = await selectedSlideId(context, stage);
    const placed = await placeOnSlide(context, slideId, textSize(payload));
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const shape = shapes.addTextBox(payload.text, placed.box);
    shape.name = `pls,fix text ${item.label}`;
    // The box hugs its text, on one line: the slide keeps the number where the
    // modeller put it instead of wrapping it into a paragraph.
    shape.textFrame.autoSizeSetting =
      PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    shape.textFrame.wordWrap = false;
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    // The text box's id is only known once this sync answers; a host that
    // swallows it never confirms one, so there is nothing here for a
    // cleanup to delete.
    await withSyncDeadline(context.sync(), "inserting the text");
    return { slideId, shapeId: shape.id, overlapping: placed.overlapping };
  });
}

// The text the source shows now, written into the box where it sits. Nothing
// here touches left, top, width, height or the font: a text link is the one
// kind whose whole look is the user's after the insert.
export async function refreshText(
  found: FoundLink,
  payload: TextPayload,
  tag: LinkTag,
): Promise<void> {
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  await PowerPoint.run(async (context) => {
    const shape = shapeAt(context, found);
    shape.textFrame.textRange.text = payload.text;
    shape.tags.add(TAG_LINK, encodeTag(tag));
    await withSyncDeadline(context.sync(), "refreshing the text");
  }).catch((error: unknown) => {
    // A box the modeller deleted since the list was drawn has the pane's own
    // sentence; wrapping the host's string here would hide it from the caller.
    if (isMissingShape(error)) throw missingShapeError(stage, error);
    const reason = error instanceof Error ? error.message : String(error);
    throw new Error(`${stage}: ${reason}`);
  });
}
