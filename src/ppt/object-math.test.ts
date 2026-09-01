import { describe, expect, it } from "vitest";
import {
  alignBoxes,
  distributeBoxes,
  matchSize,
  sameKindAndSize,
  swapBoxes,
  type ObjectBox,
} from "./object-math";

const boxes: ObjectBox[] = [
  { id: "a", left: 10, top: 20, width: 30, height: 20 },
  { id: "b", left: 60, top: 50, width: 20, height: 30 },
  { id: "c", left: 110, top: 100, width: 40, height: 10 },
];

describe("PowerPoint object geometry", () => {
  it("aligns every edge and centre against the selection bounds", () => {
    expect(alignBoxes(boxes, "left").map((move) => move.left)).toEqual([
      10, 10, 10,
    ]);
    expect(alignBoxes(boxes, "right").map((move) => move.left)).toEqual([
      120, 130, 110,
    ]);
    expect(alignBoxes(boxes, "middle").map((move) => move.top)).toEqual([
      55, 50, 60,
    ]);
  });

  it("distributes three objects while keeping the outside edges fixed", () => {
    expect(
      distributeBoxes(boxes, "horizontal").map((move) => move.left),
    ).toEqual([10, 65, 110]);
    expect(distributeBoxes(boxes, "vertical").map((move) => move.top)).toEqual([
      20, 55, 100,
    ]);
  });

  it("matches every target to the first selected size and swaps two positions", () => {
    expect(matchSize(boxes)).toEqual([
      { id: "b", width: 30, height: 20 },
      { id: "c", width: 30, height: 20 },
    ]);
    expect(swapBoxes(boxes.slice(0, 2))).toEqual([
      { id: "a", left: 60, top: 50 },
      { id: "b", left: 10, top: 20 },
    ]);
  });

  it("selects similar shapes by type and near-identical dimensions", () => {
    const source = { type: "TextBox", width: 100, height: 20 };
    expect(sameKindAndSize(source, { ...source, width: 100.5 })).toBe(true);
    expect(sameKindAndSize(source, { ...source, width: 105 })).toBe(false);
    expect(sameKindAndSize(source, { ...source, type: "Image" })).toBe(false);
  });
});
