// Occupied-box rules the free-space scan feeds on: a Mac group whose own box
// is zero still occupies the union of its children, and a dashed empty frame
// does not occupy. resolveTarget is the one reader.

import { afterEach, describe, expect, it } from "vitest";
import { overlaps } from "../layout";
import { bootPpt } from "../../test/ppt.support";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { overlapNote, OVERLAP_NOTE } from "./host";
import { resolveTarget, type InsertTarget } from "./placement";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

const SIZE = { width: 200, height: 100 };

describe("occupied boxes", () => {
  it("treats a zero-box group as occupying its children, so the next insert misses them", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      left: 36,
      top: 36,
      width: 438,
      height: 468,
    });
    const a = presentation.addShape(slide, {
      left: 502,
      top: 36,
      width: 200,
      height: 200,
    });
    const b = presentation.addShape(slide, {
      left: 502,
      top: 236,
      width: 200,
      height: 200,
    });
    presentation.groupShapes([a.id, b.id], slide.id);
    const group = slide.shapes.find((shape) => shape.type === "Group");
    expect(group).toBeDefined();
    group!.left = 0;
    group!.top = 0;
    group!.width = 0;
    group!.height = 0;

    const target: InsertTarget = { slideId: slide.id, where: "free" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    const children = { left: 502, top: 36, width: 200, height: 400 };
    expect(overlaps(resolved.placement.box, children, 12)).toBe(false);
  });

  it("does not treat a dashed empty frame as occupied", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    const frame = presentation.addShape(slide, {
      left: 36,
      top: 36,
      width: 438,
      height: 468,
    });
    Object.assign(frame, { lineDashStyle: "Dash" });

    const target: InsertTarget = { slideId: slide.id, where: "free" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    expect(resolved.placement.overlapping).toBe(false);
    // A frame that does not occupy is an empty slide: the object is centred.
    expect(resolved.placement.box).toMatchObject({
      left: 380,
      top: 220,
      width: 200,
      height: 100,
    });
  });

  it("still treats a caption as occupied when no font size is reported", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      type: "TextBox",
      text: "Revenue chart: Where = Left half",
      hasText: true,
      left: 36,
      top: 36,
      width: 438,
      height: 468,
    });

    const target: InsertTarget = { slideId: slide.id, where: "free" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    // A host that never reports a font size (mixed runs, or this fake's own
    // default) is left untrimmed rather than guessed at: the full frame
    // still counts as occupied, today's behaviour.
    expect(
      overlaps(
        resolved.placement.box,
        {
          left: 36,
          top: 36,
          width: 438,
          height: 468,
        },
        12,
      ),
    ).toBe(false);
  });

  it("trims a caption with a known font size to the lines it actually holds", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    // Full slide width, so left/right of the caption is not an option: the
    // only way this chart fits without shrinking or overlapping is below the
    // one line the caption actually holds.
    const caption = presentation.addShape(slide, {
      type: "TextBox",
      text: "Revenue chart",
      hasText: true,
      left: 36,
      top: 36,
      width: 888,
      height: 468,
    });
    Object.assign(caption, { font: { size: 11 } });

    const target: InsertTarget = { slideId: slide.id, where: "free" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    // One 11 pt line is under 25 pt tall with its insets: the chart lands in
    // the freed lower frame instead of the largest-remaining-rectangle
    // fallback (or an overlap) the untrimmed 468 pt box used to force.
    expect(resolved.placement.overlapping).toBe(false);
    expect(resolved.placement.scale).toBe(1);
    expect(resolved.placement.box.top).toBeGreaterThan(36 + 25);
  });

  // Slide 1 of the demo deck: a title, a rule and one content placeholder
  // whose text is five short numbered paragraphs plus a bold line at 14 pt -
  // the lower 60% of the 888 x 468 placeholder is empty in the real deck.
  it("lands a chart below slide 1's text instead of over it", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      type: "Placeholder",
      text: "Title",
      hasText: true,
      left: 36,
      top: 0,
      width: 888,
      height: 22,
    });
    presentation.addShape(slide, {
      left: 36,
      top: 22,
      width: 888,
      height: 2,
    });
    const body = presentation.addShape(slide, {
      type: "Placeholder",
      hasText: true,
      text: [
        `1. ${"a".repeat(40)}`,
        `2. ${"b".repeat(60)}`,
        `3. ${"c".repeat(90)}`,
        `4. ${"d".repeat(55)}`,
        `5. ${"e".repeat(70)}`,
        "Bold summary line here",
      ].join("\n"),
      left: 36,
      top: 36,
      width: 888,
      height: 468,
    });
    Object.assign(body, { font: { size: 14 } });

    const chart = { width: 900, height: 400 };
    // What the real caller (charts.ts's minPlacementScale) works out for a
    // 900 x 400 chart against MIN_SIZE 200 x 120: the higher of the two
    // ratios, 0.3.
    const minScale = 0.3;
    const target: InsertTarget = { slideId: slide.id, where: "free" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, chart, minScale),
    );
    expect(resolved.placement.overlapping).toBe(false);
    expect(resolved.placement.scale).toBeGreaterThanOrEqual(minScale);
    expect(
      overlaps(resolved.placement.box, {
        left: 36,
        top: 36,
        width: 888,
        height: 108,
      }),
    ).toBe(false);
  });
});

describe("overlapNote", () => {
  it("keeps the old sentence when nothing was free", () => {
    expect(overlapNote()).toBe(OVERLAP_NOTE);
    expect(overlapNote("whole")).toBe(OVERLAP_NOTE);
  });

  it("names the quarter that still had room", () => {
    expect(overlapNote("bottom-right")).toBe(
      "Placed over other objects: the bottom right was free at a smaller size",
    );
  });
});
