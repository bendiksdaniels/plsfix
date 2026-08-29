import { describe, expect, it } from "vitest";

import { dropBelow, overlaps, placeInFreeSpace } from "./layout";

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
});
