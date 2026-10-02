// Where an object lands on a bounded canvas without covering what is there;
// split out of layout.ts to keep both files under the line cap. Pure
// geometry. Invariant: never shrinks an object below the caller's minScale -
// past that, full size and overlapping, with the free spot's size reported.

import {
  clear,
  fitInto,
  hasArea,
  spotBox,
  type Box,
  type Canvas,
  type Placement,
  type Size,
  type Spot,
} from "./layout";

// Below this much difference, two scales count as the same one: a step of
// 0.1 landing on minScale by float error must still run that iteration.
const SCALE_EPSILON = 1e-9;

/**
 * Centred when nothing else is there; otherwise the first free spot reading
 * left to right, top to bottom, on a grid of `step`; an object too big for any spot shrinks a tenth at a time down to
 * `minScale`; past that it fits into the largest remaining rectangle when
 * that still meets `minScale`, or is centred and flagged as overlapping
 * otherwise - the hole has no area, or even the largest one would shrink the
 * object further than the caller said it may.
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
  for (
    let scale = 1;
    scale >= minScale - SCALE_EPSILON;
    scale = round1(scale - 0.1)
  ) {
    const scaled = { width: size.width * scale, height: size.height * scale };
    const box = scanGrid(scaled, occupied, canvas, margin, gap, step);
    if (box) return { box: rounded(box), scale, overlapping: false };
  }
  return fitHole(size, occupied, canvas, margin, gap, minScale);
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

// The largest remaining rectangle, fitted into, when that meets minScale; a
// hole below it is no better than no hole at all - the object goes centred
// and full size, overlapping what is there, with the hole's own size
// (however small) reported so the caller can say why.
function fitHole(
  size: Size,
  occupied: readonly Box[],
  canvas: Canvas,
  margin: number,
  gap: number,
  minScale: number,
): Placement {
  const hole = largestFreeBox(occupied, canvas, margin, gap);
  const freeSpot = hole ? namedSpotFor(hole, canvas, margin, gap) : undefined;
  if (hole) {
    const fitted = fitInto(size, hole);
    const scale = size.width > 0 ? fitted.width / size.width : 1;
    if (hasArea(fitted) && scale >= minScale - SCALE_EPSILON) {
      return { box: rounded(fitted), scale, overlapping: false, freeSpot };
    }
  }
  return {
    box: rounded(centred(size, canvas)),
    scale: 1,
    overlapping: true,
    freeSpot,
    freeSpotSize: hole
      ? { width: Math.round(hole.width), height: Math.round(hole.height) }
      : undefined,
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
