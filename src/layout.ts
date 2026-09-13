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
  /** Named spot the remaining hole sits in, when the object had to shrink or overlap. */
  freeSpot?: Spot;
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

// Where a spot sits on one axis: the content area's own start, or past the
// first half plus the gap, or (halves only) the whole axis.
type Axis = "start" | "end" | "full";

// Every spot but "whole" as its two axis choices; "whole" needs none, it IS
// the content box, so the lookup only ever covers the other six.
const SPOT_AXES: Record<Exclude<Spot, "whole">, { x: Axis; y: Axis }> = {
  "left-half": { x: "start", y: "full" },
  "right-half": { x: "end", y: "full" },
  "top-left": { x: "start", y: "start" },
  "top-right": { x: "end", y: "start" },
  "bottom-left": { x: "start", y: "end" },
  "bottom-right": { x: "end", y: "end" },
};

// One axis of a spot's box: the full side, or half of it (less the gap,
// split in two) at its start or past the gap at its end.
function axisBox(
  start: number,
  full: number,
  gap: number,
  axis: Axis,
): { pos: number; size: number } {
  if (axis === "full") return { pos: start, size: full };
  const half = (full - gap) / 2;
  return axis === "start"
    ? { pos: start, size: half }
    : { pos: start + half + gap, size: half };
}

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
  if (spot === "whole") return content;
  const { x, y } = SPOT_AXES[spot];
  const h = axisBox(content.left, content.width, gap, x);
  const v = axisBox(content.top, content.height, gap, y);
  return { left: h.pos, top: v.pos, width: h.size, height: v.size };
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

export function hasArea(box: Box): boolean {
  return box.width > 0 && box.height > 0;
}

export function unionBoxes(boxes: readonly Box[]): Box | null {
  const usable = boxes.filter(hasArea);
  if (usable.length === 0) return null;
  const left = Math.min(...usable.map((box) => box.left));
  const top = Math.min(...usable.map((box) => box.top));
  const right = Math.max(...usable.map((box) => box.left + box.width));
  const bottom = Math.max(...usable.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

// A group's own left/top/width/height when the host reports them, or the
// union of its children when that box has no area (PowerPoint for Mac, 13.09).
export function reportedBox(own: Box, children: readonly Box[]): Box {
  return hasArea(own) ? own : (unionBoxes(children) ?? own);
}

export interface FrameLook {
  fillType: string;
  hasText: boolean;
  dashStyle: string | null;
  lineVisible: boolean;
}

// The demo deck's dashed half-guides: no fill, no text, a dashed outline.
// A caption, a filled shape or a solid outline is a real object.
export function isDecorativeFrame(shape: FrameLook): boolean {
  if (shape.hasText) return false;
  if (shape.fillType !== "NoFill") return false;
  if (!shape.lineVisible) return false;
  return shape.dashStyle !== null && shape.dashStyle !== "Solid";
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
 * `minScale`; past that it fits into the largest remaining rectangle, or is
 * centred and flagged as overlapping when even that hole has no area.
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
  return fitHole(size, occupied, canvas, margin, gap);
}

const NAMED_SPOTS: Spot[] = [
  "top-left",
  "top-right",
  "bottom-left",
  "bottom-right",
  "left-half",
  "right-half",
  "whole",
];

function areaOf(box: Box): number {
  return box.width * box.height;
}

function intersection(a: Box, b: Box): Box | null {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  if (right <= left || bottom <= top) return null;
  return { left, top, width: right - left, height: bottom - top };
}

function namedSpotFor(
  box: Box,
  canvas: Canvas,
  margin: number,
  gap: number,
): Spot | undefined {
  if (!hasArea(box)) return undefined;
  let best: Spot | undefined;
  let bestOverlap = 0;
  let bestSpotArea = Infinity;
  for (const spot of NAMED_SPOTS) {
    const region = spotBox(spot, canvas, margin, gap);
    const overlap = intersection(box, region);
    const overlapArea = overlap ? areaOf(overlap) : 0;
    const spotArea = areaOf(region);
    if (
      overlapArea > bestOverlap ||
      (overlapArea === bestOverlap &&
        overlapArea > 0 &&
        spotArea < bestSpotArea)
    ) {
      best = spot;
      bestOverlap = overlapArea;
      bestSpotArea = spotArea;
    }
  }
  return best;
}

function uniqueSorted(values: number[]): number[] {
  return [...new Set(values)].sort((a, b) => a - b);
}

function largestFreeBox(
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
): Box | null {
  const left = margin;
  const top = margin;
  const right = canvas.width - margin;
  const bottom = canvas.height - margin;
  const xs = uniqueSorted([
    left,
    right,
    ...occupied.flatMap((box) => [
      box.left,
      box.left + box.width,
      box.left - gap,
      box.left + box.width + gap,
    ]),
  ]).filter((x) => x >= left && x <= right);
  const ys = uniqueSorted([
    top,
    bottom,
    ...occupied.flatMap((box) => [
      box.top,
      box.top + box.height,
      box.top - gap,
      box.top + box.height + gap,
    ]),
  ]).filter((y) => y >= top && y <= bottom);
  let best: Box | null = null;
  let bestArea = 0;
  for (let i = 0; i < xs.length; i += 1) {
    for (let j = i + 1; j < xs.length; j += 1) {
      for (let k = 0; k < ys.length; k += 1) {
        for (let l = k + 1; l < ys.length; l += 1) {
          const candidate: Box = {
            left: xs[i]!,
            top: ys[k]!,
            width: xs[j]! - xs[i]!,
            height: ys[l]! - ys[k]!,
          };
          if (!hasArea(candidate) || !clear(candidate, occupied, gap)) continue;
          const area = areaOf(candidate);
          if (area > bestArea) {
            best = candidate;
            bestArea = area;
          }
        }
      }
    }
  }
  return best;
}

function fitHole(
  size: Size,
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
): Placement {
  const hole = largestFreeBox(occupied, canvas, margin, gap);
  const freeSpot = hole ? namedSpotFor(hole, canvas, margin, gap) : undefined;
  if (hole) {
    const fitted = fitInto(size, hole);
    if (hasArea(fitted)) {
      const scale = size.width > 0 ? fitted.width / size.width : 1;
      return { box: rounded(fitted), scale, overlapping: false, freeSpot };
    }
  }
  return {
    box: rounded(centred(size, canvas)),
    scale: 1,
    overlapping: true,
    freeSpot,
  };
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
