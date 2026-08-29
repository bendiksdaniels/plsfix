// Where a new object goes so it covers nothing: beside an anchor on a sheet
// that has no edge (a chart the add-in inserts), or in the free space of a
// bounded canvas (a slide). Pure geometry in the caller's units; the callers
// read the occupied boxes and write the one that comes back.

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

/** Right of the anchor first, then below it, then below whatever is in the way. */
export function placeBeside(
  anchor: Box,
  size: Size,
  occupied: readonly Box[],
  gap: number,
): Box {
  const right = {
    ...size,
    left: anchor.left + anchor.width + gap,
    top: anchor.top,
  };
  if (clear(right, occupied, gap)) return right;
  const below = {
    ...size,
    left: anchor.left,
    top: anchor.top + anchor.height + gap,
  };
  if (clear(below, occupied, gap)) return below;
  const rightColumn = descend(right, occupied, gap);
  const leftColumn = descend(below, occupied, gap);
  return leftColumn.top < rightColumn.top ? leftColumn : rightColumn;
}

// Every step moves the box under at least one of the boxes blocking it, so
// the walk ends within one step per occupied box.
function descend(start: Box, occupied: readonly Box[], gap: number): Box {
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
 * The first free spot reading left to right, top to bottom, on a grid of
 * `step`; an object too big for any spot shrinks a tenth at a time down to
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
  for (let scale = 1; scale >= minScale - 1e-9; scale = round1(scale - 0.1)) {
    const scaled = { width: size.width * scale, height: size.height * scale };
    const box = scanGrid(scaled, occupied, canvas, margin, gap, step);
    if (box) return { box: rounded(box), scale, overlapping: false };
  }
  const centred = {
    ...size,
    left: (canvas.width - size.width) / 2,
    top: (canvas.height - size.height) / 2,
  };
  return { box: rounded(centred), scale: 1, overlapping: true };
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
