import { describe, expect, it } from "vitest";

import { SLIDE_16_9 } from "./link/status";
import {
  dropBelow,
  fitInto,
  overlaps,
  placeInFreeSpace,
  spotBox,
  type Spot,
} from "./layout";

const size = { width: 300, height: 200 };

describe("dropBelow", () => {
  const box = { ...size, left: 100, top: 100 };

  it("leaves a box nothing is in the way of alone", () => {
    expect(dropBelow(box, [], 10)).toEqual(box);
  });

  it("steps under everything in the way", () => {
    const first = { left: 100, top: 100, width: 300, height: 200 };
    const second = { left: 100, top: 320, width: 300, height: 400 };
    const dropped = dropBelow(box, [first, second], 10);
    expect(dropped).toMatchObject({ left: 100, top: 730 });
    expect(overlaps(dropped, first)).toBe(false);
    expect(overlaps(dropped, second)).toBe(false);
  });
});

describe("placeInFreeSpace", () => {
  const slide = { width: 960, height: 540 };

  it("centres the only object on an empty slide", () => {
    expect(placeInFreeSpace(size, [], slide, 36, 12)).toEqual({
      box: { left: 330, top: 170, width: 300, height: 200 },
      scale: 1,
      overlapping: false,
    });
  });

  it("skips the objects already there", () => {
    const taken = { left: 36, top: 36, width: 500, height: 200 };
    const placement = placeInFreeSpace(size, [taken], slide, 36, 12);
    expect(placement.overlapping).toBe(false);
    expect(overlaps(placement.box, taken, 12)).toBe(false);
    expect(placement.box).toMatchObject({ left: 548, top: 36 });
  });

  it("shrinks when only a smaller spot is free, and says so", () => {
    const wall = { left: 36, top: 36, width: 888, height: 300 };
    const placement = placeInFreeSpace(size, [wall], slide, 36, 12);
    expect(placement.scale).toBeLessThan(1);
    expect(placement.overlapping).toBe(false);
    expect(placement.box.height).toBeLessThanOrEqual(540 - 36 - 348);
  });

  it("centres and flags an object that fits nowhere", () => {
    const everything = { left: 0, top: 0, width: 960, height: 540 };
    const placement = placeInFreeSpace(size, [everything], slide, 36, 12);
    expect(placement.overlapping).toBe(true);
    expect(placement.box).toMatchObject({ left: 330, top: 170 });
  });

  // scanGrid's own early return: the object is too big for the canvas even
  // at minScale, before a single grid cell is even tried - not the "every
  // cell is blocked" path the test above covers.
  it("centres and flags an object too big for the canvas at any scale", () => {
    const huge = { width: 3000, height: 3000 };
    const speck = { left: 0, top: 0, width: 1, height: 1 };
    const placement = placeInFreeSpace(huge, [speck], slide, 36, 12);
    expect(placement.overlapping).toBe(true);
    expect(placement.scale).toBe(1);
    expect(placement.box).toMatchObject({
      left: (960 - 3000) / 2,
      top: (540 - 3000) / 2,
    });
  });
});

describe("spotBox", () => {
  const margin = 36;
  const gap = 12;
  // The 16:9 slide minus the margin on every side: 960-72 x 540-72.
  const content = { left: 36, top: 36, width: 888, height: 468 };
  const half = { width: 438, height: 468 };
  const quarter = { width: 438, height: 228 };

  it.each<[Spot, ReturnType<typeof spotBox>]>([
    ["whole", content],
    ["left-half", { left: 36, top: 36, ...half }],
    ["right-half", { left: 486, top: 36, ...half }],
    ["top-left", { left: 36, top: 36, ...quarter }],
    ["top-right", { left: 486, top: 36, ...quarter }],
    ["bottom-left", { left: 36, top: 276, ...quarter }],
    ["bottom-right", { left: 486, top: 276, ...quarter }],
  ])("%s", (spot, box) => {
    expect(spotBox(spot, SLIDE_16_9, margin, gap)).toEqual(box);
  });

  it("keeps the gap between the two halves and all four quarters", () => {
    const left = spotBox("left-half", SLIDE_16_9, margin, gap);
    const right = spotBox("right-half", SLIDE_16_9, margin, gap);
    expect(right.left - (left.left + left.width)).toBe(gap);

    const topLeft = spotBox("top-left", SLIDE_16_9, margin, gap);
    const bottomLeft = spotBox("bottom-left", SLIDE_16_9, margin, gap);
    expect(bottomLeft.top - (topLeft.top + topLeft.height)).toBe(gap);
  });
});

describe("fitInto", () => {
  const box = { left: 100, top: 50, width: 400, height: 200 };

  it("binds to the box's width when the size is proportionally wider", () => {
    expect(fitInto({ width: 800, height: 100 }, box)).toEqual({
      left: 100,
      top: 125,
      width: 400,
      height: 50,
    });
  });

  it("binds to the box's height when the size is proportionally taller", () => {
    expect(fitInto({ width: 100, height: 400 }, box)).toEqual({
      left: 275,
      top: 50,
      width: 50,
      height: 200,
    });
  });

  it("returns the box unchanged when the size already matches it exactly", () => {
    expect(fitInto({ width: 400, height: 200 }, box)).toEqual(box);
  });

  // Placement.scale is documented as never above 1 ("smaller when it had to
  // shrink to fit"): a size already smaller than the box stays its own size,
  // centred, rather than being blown up to fill the spot.
  it("never enlarges a size that already fits", () => {
    expect(fitInto({ width: 40, height: 20 }, box)).toEqual({
      left: 280,
      top: 140,
      width: 40,
      height: 20,
    });
  });
});
