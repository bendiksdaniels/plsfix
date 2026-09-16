// The box/canvas types every placement caller shares, and the geometry that
// does not need a search: a named spot's own rectangle, fitting a size into a
// box keeping its aspect, whether two boxes overlap, a group's box from its
// children, a decorative frame's look, and dropping a box straight down clear
// of what is in the way (src/excel/chart-place.ts). The free-space search
// itself - placeInFreeSpace - is free-space.ts, split out to keep both files
// under the line cap. Pure geometry in the caller's units.

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
  /** The largest free box's own size, set when overlapping and a hole (too
   * small for minScale) was found anyway: what the toast tells the user to
   * explain the overlap ("the largest free spot is 888 x 40 pt"). */
  freeSpotSize?: Size;
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

// Exported for free-space.ts's own scan and largest-hole search: the same
// "nothing occupied is in the way, gap included" test dropBelow uses here.
export function clear(
  box: Box,
  occupied: readonly Box[],
  gap: number,
): boolean {
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
