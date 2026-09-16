import { describe, expect, it } from "vitest";

import { placeInFreeSpace } from "./free-space";
import { SLIDE_16_9 } from "./link/status";
import {
  dropBelow,
  fitInto,
  hasArea,
  isDecorativeFrame,
  overlaps,
  reportedBox,
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

  // A floor, not a suggestion (chart-place): scanGrid's own early return (the
  // object is too big for the canvas even at minScale) used to fall back to
  // fitHole shrinking it into the remaining content regardless of minScale -
  // a chart on a busy slide could land as a sliver nobody could read. Past
  // minScale, the answer is the pre-v2.8.4 one again: centred, full size,
  // overlapping, with the hole it could not use reported so the caller can
  // say why.
  it("overlaps full-size, past minScale, rather than shrinking into a small hole", () => {
    const huge = { width: 3000, height: 3000 };
    const speck = { left: 0, top: 0, width: 1, height: 1 };
    const placement = placeInFreeSpace(huge, [speck], slide, 36, 12);
    expect(placement.overlapping).toBe(true);
    expect(placement.scale).toBe(1);
    expect(placement.box).toMatchObject({ left: -1020, top: -1230 });
    // The whole content area (888 x 468) was free, just not big enough.
    expect(placement.freeSpotSize).toEqual({ width: 888, height: 468 });
  });

  it("still shrinks into a hole that clears minScale, exactly as before", () => {
    const huge = { width: 3000, height: 3000 };
    const speck = { left: 0, top: 0, width: 1, height: 1 };
    const placement = placeInFreeSpace(huge, [speck], slide, 36, 12, 0.1);
    expect(placement.overlapping).toBe(false);
    expect(placement.scale).toBeLessThan(1);
    expect(placement.box.width).toBeLessThanOrEqual(888);
    expect(placement.box.height).toBeLessThanOrEqual(468);
  });

  // minScale 1 forbids the 0.1 grid shrink, so this always fell to fitHole.
  // The strip between the walls only fits the object at 0.72: below the
  // floor the caller asked for, so this is now the overlap, not the shrink.
  it("overlaps full-size when the only free strip is narrower than minScale allows", () => {
    const left = { left: 0, top: 0, width: 41, height: 540 };
    const right = { left: 280, top: 0, width: 680, height: 540 };
    const placement = placeInFreeSpace(size, [left, right], slide, 36, 12, 1);
    expect(placement.overlapping).toBe(true);
    expect(placement.scale).toBe(1);
    expect(placement.box).toMatchObject({ left: 330, top: 170 });
    expect(placement.freeSpot).toBe("left-half");
    expect(placement.freeSpotSize).toEqual({ width: 215, height: 468 });
  });

  it("still fits that same strip once minScale allows the shrink", () => {
    const left = { left: 0, top: 0, width: 41, height: 540 };
    const right = { left: 280, top: 0, width: 680, height: 540 };
    const placement = placeInFreeSpace(size, [left, right], slide, 36, 12, 0.5);
    expect(placement.overlapping).toBe(false);
    expect(placement.scale).toBeLessThan(1);
    expect(overlaps(placement.box, left, 12)).toBe(false);
    expect(overlaps(placement.box, right, 12)).toBe(false);
  });

  it("overlaps full-size over a small remaining hole, naming its size and spot", () => {
    const hole = { left: 700, top: 400, width: 80, height: 60 };
    const blocked = [
      { left: 0, top: 0, width: 960, height: 400 },
      { left: 0, top: 400, width: 700, height: 140 },
      { left: 780, top: 400, width: 180, height: 140 },
      { left: 700, top: 460, width: 80, height: 80 },
    ];
    expect(overlaps(hole, blocked[0]!)).toBe(false);
    const placement = placeInFreeSpace(size, blocked, slide, 36, 12, 1);
    expect(placement.overlapping).toBe(true);
    expect(placement.scale).toBe(1);
    expect(placement.box).toMatchObject({ left: 330, top: 170 });
    expect(placement.freeSpot).toBe("bottom-right");
    expect(placement.freeSpotSize).toEqual({ width: 56, height: 36 });
  });

  // Slide 3 of the demo deck: a title, a rule, two chart groups already
  // sitting in their dashed frames (excluded by placement.ts before this
  // reaches placeInFreeSpace) and two captions. The only free space left is
  // an 888 x 132 strip above the charts - too short at minScale 0.4 for a
  // third chart's natural 500 x 400 - so a Free-space insert here must
  // overlap rather than land as an unreadable sliver.
  it("slide 3: overlaps a busy slide rather than shrinking a chart into the top strip", () => {
    const title = { left: 36, top: 0, width: 888, height: 22 };
    const rule = { left: 36, top: 22, width: 888, height: 2 };
    const columnGroup = { left: 60, top: 200, width: 390, height: 200 };
    const pieGroup = { left: 500, top: 180, width: 400, height: 240 };
    const caption1 = { left: 36, top: 508, width: 438, height: 24 };
    const caption2 = { left: 486, top: 508, width: 438, height: 24 };
    const occupied = [title, rule, columnGroup, pieGroup, caption1, caption2];
    const chart = { width: 500, height: 400 };
    const placement = placeInFreeSpace(chart, occupied, slide, 36, 12, 0.4);
    expect(placement.overlapping).toBe(true);
    expect(placement.scale).toBe(1);
    expect(placement.freeSpotSize).toEqual({ width: 888, height: 132 });
    expect(placement.box).toMatchObject({ width: 500, height: 400 });
  });
});

describe("reportedBox", () => {
  const children = [
    { left: 400, top: 40, width: 200, height: 120 },
    { left: 400, top: 160, width: 200, height: 120 },
  ];

  it("keeps a box that has area", () => {
    const own = { left: 10, top: 10, width: 50, height: 50 };
    expect(reportedBox(own, children)).toEqual(own);
  });

  it("unions the children when the host reports a zero box", () => {
    expect(
      reportedBox({ left: 0, top: 0, width: 0, height: 0 }, children),
    ).toEqual({ left: 400, top: 40, width: 200, height: 240 });
  });

  it("treats a zero-area box as empty", () => {
    expect(hasArea({ left: 0, top: 0, width: 0, height: 0 })).toBe(false);
    expect(hasArea({ left: 10, top: 10, width: 1, height: 1 })).toBe(true);
  });
});

describe("isDecorativeFrame", () => {
  it("is a dashed, empty, unfilled rectangle", () => {
    expect(
      isDecorativeFrame({
        fillType: "NoFill",
        hasText: false,
        dashStyle: "Dash",
        lineVisible: true,
      }),
    ).toBe(true);
  });

  it("is not a caption, a filled shape, or a solid outline", () => {
    const frame = {
      fillType: "NoFill",
      hasText: false,
      dashStyle: "Dash",
      lineVisible: true,
    };
    expect(isDecorativeFrame({ ...frame, hasText: true })).toBe(false);
    expect(isDecorativeFrame({ ...frame, fillType: "Solid" })).toBe(false);
    expect(isDecorativeFrame({ ...frame, dashStyle: "Solid" })).toBe(false);
    expect(isDecorativeFrame({ ...frame, lineVisible: false })).toBe(false);
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
