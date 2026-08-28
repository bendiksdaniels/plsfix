// Link state rules: relay revision vs. tag revision decides current/stale/
// missing/wrongKey, and the pixel-to-point maths that fits a pushed PNG onto
// a slide, capped to the margins and centred. No Office.js - callers pass in
// plain numbers and get plain numbers back.

import type { Source } from "./model";

export type LinkStatus = "current" | "updateAvailable" | "missing" | "wrongKey";

export interface RelayStatus {
  id: string;
  rev: number | null;
  pushedAt: number | null;
  error?: "auth";
}

export function deriveStatus(
  tagRev: number,
  relay: RelayStatus | undefined,
): LinkStatus {
  if (relay === undefined) return "missing";
  if (relay.error === "auth") return "wrongKey";
  if (relay.rev === null) return "missing";
  // Any inequality, not just a higher rev: a link swept at its 7-day TTL comes
  // back from the next push as rev 1, so a relay rev *below* the tag's is the
  // ordinary "the deck is stale" case, not an impossibility.
  if (relay.rev !== tagRev) return "updateAvailable";
  return "current";
}

export function sourceChanged(before: Source, after: Source): boolean {
  return before.workbook.toLowerCase() !== after.workbook.toLowerCase();
}

const DEFAULT_ASPECT_TOLERANCE = 0.005;

export function aspectChanged(
  width: number,
  height: number,
  pngWidth: number,
  pngHeight: number,
  tolerance: number = DEFAULT_ASPECT_TOLERANCE,
): boolean {
  const current = width / height;
  const source = pngWidth / pngHeight;
  return Math.abs(current - source) / source > tolerance;
}

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export const SLIDE_16_9 = { width: 960, height: 540 }; // points

const PIXELS_TO_POINTS = 0.75; // 72pt / 96dpi
const DEFAULT_SLIDE_MARGIN = 36; // half an inch, in points

export function fitToSlide(
  pngWidth: number,
  pngHeight: number,
  slide: { width: number; height: number } = SLIDE_16_9,
  margin: number = DEFAULT_SLIDE_MARGIN,
): Box {
  const rawWidth = pngWidth * PIXELS_TO_POINTS;
  const rawHeight = pngHeight * PIXELS_TO_POINTS;
  const maxWidth = slide.width - 2 * margin;
  const maxHeight = slide.height - 2 * margin;
  const scale = Math.min(1, maxWidth / rawWidth, maxHeight / rawHeight);

  const width = Math.round(rawWidth * scale);
  const height = Math.round(rawHeight * scale);
  return {
    left: Math.round((slide.width - width) / 2),
    top: Math.round((slide.height - height) / 2),
    width,
    height,
  };
}
