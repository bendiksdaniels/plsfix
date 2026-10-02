// How much of a text-bearing shape's own box its current text needs, so the
// free-space search sees only that and not a placeholder's whole frame. Pure:
// no Office.js, only chart-shapes.ts's own text-width estimate.
// Invariant: the answer never exceeds the frame's own height.

import { textWidth } from "./chart-shapes";

// PowerPoint's own default text-frame insets, in points: 0.05 in top and
// bottom, 0.1 in left and right (the same left/right figure textWidth's own
// LABEL_PAD covers for a label's box).
const TOP_INSET = 3.6;
const BOTTOM_INSET = 3.6;
const SIDE_INSET = 14.4;
// A single-spaced line's height as a multiple of its font size.
const LINE_HEIGHT_FACTOR = 1.2;

function paragraphs(text: string): string[] {
  return text.split(/\r\n|[\r\n]/);
}

// How many lines one paragraph wraps to at fontSize inside innerWidth: at
// least one, more once textWidth's estimate outgrows the room, the same rule
// PowerPoint's own word wrap would apply.
function linesFor(
  paragraph: string,
  fontSize: number,
  innerWidth: number,
): number {
  if (innerWidth <= 0) return 1;
  return Math.max(1, Math.ceil(textWidth(paragraph, fontSize) / innerWidth));
}

// The vertical room `text` needs inside a frame of frameWidth x frameHeight
// at fontSize: the top and bottom insets plus one line per LINE_HEIGHT_FACTOR
// x fontSize, wrapped paragraph by paragraph over the width left once the
// side insets are taken out - never more than the frame itself, since a
// bigger estimate would grow a shape's own occupied box rather than shrink
// it.
export function textOccupiedHeight(
  text: string,
  fontSize: number,
  frameWidth: number,
  frameHeight: number,
): number {
  const innerWidth = frameWidth - SIDE_INSET;
  const lines = paragraphs(text).reduce(
    (total, paragraph) => total + linesFor(paragraph, fontSize, innerWidth),
    0,
  );
  const height =
    TOP_INSET + lines * LINE_HEIGHT_FACTOR * fontSize + BOTTOM_INSET;
  return Math.min(height, frameHeight);
}
