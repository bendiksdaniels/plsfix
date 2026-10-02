// Line count, wrap and the frame clamp textOccupiedHeight builds an occupied
// box's height from: insets, one line per LINE_HEIGHT_FACTOR x fontSize per
// paragraph, wrapped by textWidth, never past the frame itself.

import { describe, expect, it } from "vitest";
import { textOccupiedHeight } from "./text-extent";

describe("textOccupiedHeight", () => {
  it("is the two insets plus one line for a single short line", () => {
    // 3.6 top + 1 x 1.2 x 11 + 3.6 bottom.
    expect(textOccupiedHeight("Revenue chart", 11, 438, 24)).toBeCloseTo(
      3.6 + 1.2 * 11 + 3.6,
      6,
    );
  });

  it("adds a line per paragraph break", () => {
    const oneLine = textOccupiedHeight("One short line", 14, 888, 468);
    const threeLines = textOccupiedHeight(
      "One short line\nAnother\nA third",
      14,
      888,
      468,
    );
    expect(threeLines).toBeCloseTo(oneLine + 2 * 1.2 * 14, 6);
  });

  it("wraps a paragraph wider than the frame into more than one line", () => {
    const narrow = textOccupiedHeight(
      "A very long single paragraph",
      14,
      60,
      468,
    );
    const wide = textOccupiedHeight(
      "A very long single paragraph",
      14,
      900,
      468,
    );
    expect(narrow).toBeGreaterThan(wide);
  });

  it("keeps five numbered lines plus a bold line well under a full placeholder", () => {
    // Slide 1 of the demo deck: five 40-90 character numbered paragraphs plus
    // one bold line, 14 pt, in an 888 x 468 content placeholder - the lower
    // 60% of it is empty in the real deck.
    const paragraphs = [
      `1. ${"a".repeat(40)}`,
      `2. ${"b".repeat(60)}`,
      `3. ${"c".repeat(90)}`,
      `4. ${"d".repeat(55)}`,
      `5. ${"e".repeat(70)}`,
      "Bold summary line here",
    ];
    const height = textOccupiedHeight(paragraphs.join("\n"), 14, 888, 468);
    expect(height).toBeGreaterThan(0);
    expect(height).toBeLessThan(468 * 0.5);
  });

  it("never exceeds the frame's own height", () => {
    const height = textOccupiedHeight("word ".repeat(200), 14, 100, 50);
    expect(height).toBe(50);
  });

  it("accepts \\r\\n and bare \\r paragraph breaks the same as \\n", () => {
    const lf = textOccupiedHeight("A\nB\nC", 12, 400, 400);
    const crlf = textOccupiedHeight("A\r\nB\r\nC", 12, 400, 400);
    const cr = textOccupiedHeight("A\rB\rC", 12, 400, 400);
    expect(crlf).toBeCloseTo(lf, 6);
    expect(cr).toBeCloseTo(lf, 6);
  });
});
