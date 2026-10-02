// A1 geometry: what an address means as a rectangle, and which edits count as
// landing inside a linked range. The unreadable cases matter most - a missed
// overlap is a stale picture nobody is told about.

import { describe, expect, it } from "vitest";
import { intersects, overlaps, parseAreas, parseRect } from "./geometry";

describe("parseRect", () => {
  it("reads a cell, a block and an absolute block the same way", () => {
    expect(parseRect("B4")).toEqual({ top: 4, left: 2, bottom: 4, right: 2 });
    expect(parseRect("B4:F5")).toEqual({
      top: 4,
      left: 2,
      bottom: 5,
      right: 6,
    });
    expect(parseRect("$B$4:$F$5")).toEqual(parseRect("B4:F5"));
  });

  it("normalises corners given the other way round", () => {
    expect(parseRect("F5:B4")).toEqual(parseRect("B4:F5"));
  });

  it("counts columns past Z", () => {
    expect(parseRect("AA1")?.left).toBe(27);
    expect(parseRect("XFD1")?.left).toBe(16384);
  });

  it("gives whole columns and whole rows the grid's own size", () => {
    expect(parseRect("A:C")).toEqual({
      top: 1,
      left: 1,
      bottom: 1048576,
      right: 3,
    });
    expect(parseRect("3:5")).toEqual({
      top: 3,
      left: 1,
      bottom: 5,
      right: 16384,
    });
  });

  it("refuses what is not an address", () => {
    for (const text of ["", "B", "4", "$", "A1:B", "A1:B2:C3", "#REF!"]) {
      expect(parseRect(text)).toBeNull();
    }
  });
});

describe("parseAreas", () => {
  it("reads a multi-area address", () => {
    expect(parseAreas("A1:B2,D4:E5")).toHaveLength(2);
  });

  // All or nothing: half a read address would answer "no overlap" for the half
  // it could not see.
  it("is null when any one area cannot be read", () => {
    expect(parseAreas("A1:B2,#REF!")).toBeNull();
  });
});

describe("overlaps", () => {
  it("counts a shared edge as an overlap", () => {
    const block = { top: 4, left: 2, bottom: 5, right: 6 };
    expect(overlaps(block, { top: 5, left: 6, bottom: 9, right: 9 })).toBe(
      true,
    );
    expect(overlaps(block, { top: 6, left: 7, bottom: 9, right: 9 })).toBe(
      false,
    );
  });
});

describe("intersects", () => {
  it("sees an edit inside, on the edge and outside an anchored block", () => {
    expect(intersects("$B$4:$F$5", "C5")).toBe(true);
    expect(intersects("$B$4:$F$5", "B4")).toBe(true);
    expect(intersects("$B$4:$F$5", "G5")).toBe(false);
    expect(intersects("$B$4:$F$5", "B6:F9")).toBe(false);
  });

  it("sees a whole-column or whole-row edit that crosses the block", () => {
    expect(intersects("B4:F5", "C:C")).toBe(true);
    expect(intersects("B4:F5", "A:A")).toBe(false);
    expect(intersects("B4:F5", "5:5")).toBe(true);
    expect(intersects("B4:F5", "9:9")).toBe(false);
  });

  it("sees a multi-area paste that touches the block with one of its areas", () => {
    expect(intersects("B4:F5", "H1:H2,C4")).toBe(true);
    expect(intersects("B4:F5", "H1:H2,C9")).toBe(false);
  });

  // A push too many costs a render; a push too few leaves a deck showing last
  // week's numbers with nothing on screen to say so.
  it("treats an address it cannot read as a hit", () => {
    expect(intersects("#REF!", "C5")).toBe(true);
    expect(intersects("B4:F5", "")).toBe(true);
  });
});
