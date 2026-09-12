// The three hygiene ladders of src/cycles.ts - indent, horizontal alignment
// and the font underline - in their own file: cycles.test.ts already stands at
// the 400-line ceiling. What is proved here is the matching rule, because
// Excel rewrites every one of these on read-back (lessons 2026-08-27).

import { describe, expect, it } from "vitest";
import {
  ALIGN_CYCLE,
  canonicalAlignment,
  canonicalUnderline,
  INDENT_CYCLE,
  nextAlignment,
  nextIndent,
  nextUnderline,
  UNDERLINE_CYCLE,
} from "./cycles";

describe("hygiene cycles", () => {
  it("lists the three ladders the hygiene buttons step", () => {
    expect(INDENT_CYCLE).toEqual([0, 1, 2, 3]);
    expect(ALIGN_CYCLE).toEqual(["Left", "Center", "Right", "General"]);
    expect(UNDERLINE_CYCLE).toEqual(["Single", "Double", "None"]);
  });

  it("steps the indent one level at a time and wraps home", () => {
    expect([0, 1, 2, 3].map((level) => nextIndent(level))).toEqual([
      1, 2, 3, 0,
    ]);
  });

  it("reads a null indent, and one nobody set, as no indent", () => {
    expect(nextIndent(null)).toBe(1);
    expect(nextIndent(undefined)).toBe(1);
    // A hand-set level outside the ladder steps to its foot, the way the other
    // cycles treat a look we did not apply.
    expect(nextIndent(7)).toBe(0);
    expect(nextIndent(-2)).toBe(1);
  });

  it("steps the alignment Left, Center, Right, then back to General", () => {
    expect(
      ["Left", "Center", "Right", "General"].map((at) => nextAlignment(at)),
    ).toEqual(["Center", "Right", "General", "Left"]);
  });

  it("compares the alignment Excel gives back without minding its casing", () => {
    expect(nextAlignment("left")).toBe("Center");
    expect(nextAlignment(" CENTER ")).toBe("Right");
    // A range whose cells disagree reports nothing, and an alignment we do not
    // cycle (Fill, Justify) is nobody's rung: both start at Left.
    expect(nextAlignment(null)).toBe("Left");
    expect(nextAlignment("Justify")).toBe("Left");
  });

  it("steps the underline Single, Double, then off", () => {
    expect(["Single", "Double", "None"].map((at) => nextUnderline(at))).toEqual(
      ["Double", "None", "Single"],
    );
  });

  it("counts an accounting underline as the plain one it draws", () => {
    // Excel answers SingleAccountant for the accounting underline; it is the
    // same rung of the cycle (lessons 2026-08-27).
    expect(nextUnderline("SingleAccountant")).toBe("Double");
    expect(nextUnderline("DoubleAccountant")).toBe("None");
    expect(nextUnderline(null)).toBe("Single");
    expect(canonicalUnderline("SingleAccountant")).toBe("single");
    expect(canonicalAlignment(" Center ")).toBe("center");
  });
});
