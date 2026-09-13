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

  it("still treats a caption as occupied", async () => {
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
