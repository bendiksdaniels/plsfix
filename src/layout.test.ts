import { describe, expect, it } from "vitest";

import { overlaps, placeBeside, placeInFreeSpace } from "./layout";

const anchor = { left: 100, top: 100, width: 200, height: 120 };
const size = { width: 300, height: 200 };

describe("placeBeside", () => {
  it("goes right of the anchor when that is free", () => {
    expect(placeBeside(anchor, size, [], 10)).toEqual({
      left: 310,
      top: 100,
      width: 300,
      height: 200,
    });
  });

  it("goes below the anchor when the right is taken", () => {
    const right = { left: 320, top: 100, width: 300, height: 120 };
    expect(placeBeside(anchor, size, [right], 10)).toMatchObject({
      left: 100,
      top: 230,
    });
  });

  it("steps under everything in the way and takes the higher column", () => {
    const right = { left: 310, top: 100, width: 300, height: 200 };
    const below = { left: 100, top: 230, width: 300, height: 400 };
    const box = placeBeside(anchor, size, [right, below], 10);
    expect(box).toMatchObject({ left: 310, top: 640 });
    expect(overlaps(box, right)).toBe(false);
    expect(overlaps(box, below)).toBe(false);
  });
});

describe("placeInFreeSpace", () => {
  const slide = { width: 960, height: 540 };

  it("starts at the top-left margin of an empty slide", () => {
    expect(placeInFreeSpace(size, [], slide, 36, 12)).toEqual({
      box: { left: 36, top: 36, width: 300, height: 200 },
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
});
