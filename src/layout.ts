// Where a new object goes so it covers nothing: dropped clear of what is in the
// way (a chart the add-in inserts, positioned by src/excel/chart-place.ts), or
// in the free space of a bounded canvas (a slide). Pure geometry in the
// caller's units; the callers read the occupied boxes and write what comes back.

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface Canvas {
  width: number;
  height: number;
}

export interface Placement {
  box: Box;
  /** 1 when the object kept its size; smaller when it had to shrink to fit. */
  scale: number;
  /** True when nothing fit and the object sits centred over whatever is there. */
  overlapping: boolean;
}

// A named half or quarter of the canvas, or the whole content area: what the
// PowerPoint inbox's "Where" picker offers beside "Free space" and "Selected
// shape" (those two need Office.js, so they live in src/ppt/placement.ts).
export type Spot =
  | "left-half"
  | "right-half"
  | "top-left"
  | "top-right"
  | "bottom-left"
  | "bottom-right"
  | "whole";

// The named box a spot occupies: the canvas minus margin on every side
// ("whole"), or that content area split into two or four with gap between
// the pieces - the same margin and gap placeInFreeSpace keeps.
export function spotBox(
  spot: Spot,
  canvas: Canvas,
  margin: number,
  gap: number,
): Box {
  const content: Box = {
    left: margin,
    top: margin,
    width: canvas.width - 2 * margin,
    height: canvas.height - 2 * margin,
  };
  const halfWidth = (content.width - gap) / 2;
  const halfHeight = (content.height - gap) / 2;
  const leftX = content.left;
  const rightX = content.left + halfWidth + gap;
  const topY = content.top;
  const bottomY = content.top + halfHeight + gap;
  switch (spot) {
    case "whole":
      return content;
    case "left-half":
      return {
        left: leftX,
        top: topY,
        width: halfWidth,
        height: content.height,
      };
    case "right-half":
      return {
        left: rightX,
        top: topY,
        width: halfWidth,
        height: content.height,
      };
    case "top-left":
      return { left: leftX, top: topY, width: halfWidth, height: halfHeight };
    case "top-right":
      return { left: rightX, top: topY, width: halfWidth, height: halfHeight };
    case "bottom-left":
      return {
        left: leftX,
        top: bottomY,
        width: halfWidth,
        height: halfHeight,
      };
    case "bottom-right":
      return {
        left: rightX,
        top: bottomY,
        width: halfWidth,
        height: halfHeight,
      };
  }
}

// size scaled down (never up: Placement.scale is never above 1) to fit inside
// box, aspect kept, then centred in it. Used wherever a link's native size
// has to land inside a spot or a selected shape's box instead of free space.
export function fitInto(size: Size, box: Box): Box {
  const scale = Math.min(1, box.width / size.width, box.height / size.height);
  const width = size.width * scale;
  const height = size.height * scale;
  return {
    left: box.left + (box.width - width) / 2,
    top: box.top + (box.height - height) / 2,
    width,
    height,
  };
}

export function overlaps(a: Box, b: Box, gap = 0): boolean {
  return (
    a.left < b.left + b.width + gap &&
    b.left < a.left + a.width + gap &&
    a.top < b.top + b.height + gap &&
    b.top < a.top + a.height + gap
  );
}

function clear(box: Box, occupied: readonly Box[], gap: number): boolean {
  return occupied.every((other) => !overlaps(box, other, gap));
}

/**
 * The same box, moved straight down until nothing occupied is in its way.
 * Every step moves it under at least one of the boxes blocking it, so the walk
 * ends within one step per occupied box.
 */
export function dropBelow(
  start: Box,
  occupied: readonly Box[],
  gap: number,
): Box {
  let box = start;
  for (let step = 0; step <= occupied.length; step += 1) {
    const blocking = occupied.filter((other) => overlaps(box, other, gap));
    if (blocking.length === 0) return box;
    const bottom = Math.max(
      ...blocking.map((other) => other.top + other.height),
    );
    box = { ...box, top: bottom + gap };
  }
  return box;
}

/**
 * Centred when nothing else is there; otherwise the first free spot reading
 * left to right, top to bottom, on a grid of `step`; an object too big for any spot shrinks a tenth at a time down to
 * `minScale`; past that it is centred and flagged as overlapping.
 */
export function placeInFreeSpace(
  size: Size,
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
  minScale = 0.4,
  step = 8,
): Placement {
  if (occupied.length === 0) {
    return {
      box: rounded(centred(size, canvas)),
      scale: 1,
      overlapping: false,
    };
  }
  for (let scale = 1; scale >= minScale - 1e-9; scale = round1(scale - 0.1)) {
    const scaled = { width: size.width * scale, height: size.height * scale };
    const box = scanGrid(scaled, occupied, canvas, margin, gap, step);
    if (box) return { box: rounded(box), scale, overlapping: false };
  }
  return { box: rounded(centred(size, canvas)), scale: 1, overlapping: true };
}

// Alone on the canvas an object is centred, as a slide reads best; the
// top-left scan is for company.
function centred(size: Size, canvas: Canvas): Box {
  return {
    ...size,
    left: (canvas.width - size.width) / 2,
    top: (canvas.height - size.height) / 2,
  };
}

function scanGrid(
  size: Size,
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
  step: number,
): Box | null {
  const maxLeft = canvas.width - margin - size.width;
  const maxTop = canvas.height - margin - size.height;
  if (maxLeft < margin || maxTop < margin) return null;
  for (let top = margin; top <= maxTop; top += step) {
    for (let left = margin; left <= maxLeft; left += step) {
      const box = { ...size, left, top };
      if (clear(box, occupied, gap)) return box;
    }
  }
  return null;
}

function round1(value: number): number {
  return Math.round(value * 10) / 10;
}

function rounded(box: Box): Box {
  return {
    left: Math.round(box.left),
    top: Math.round(box.top),
    width: Math.round(box.width),
    height: Math.round(box.height),
  };
}
