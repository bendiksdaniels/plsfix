// The Slide/Where pickers' target resolution: src/ppt/placement.ts's
// resolveTarget and finishTarget directly, then every insert path (picture,
// table, text, chart) end to end through links.insertFromInbox. Kept out of
// links.audit.test.ts, already at the 400-line cap.

import { afterEach, describe, expect, it } from "vitest";
import { spotBox, type Spot } from "../layout";
import { createWorkspace } from "../link/workspace";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "../../test/ppt.support";
import { uninstallFakePpt, type FakePptShape } from "../../test/fakeppt";
import { fakePng } from "../../test/fakepng";
import {
  finishTarget,
  resolveTarget,
  SLIDE,
  type InsertTarget,
} from "./placement";

afterEach(() => {
  uninstallFakePpt();
});

const SIZE = { width: 200, height: 100 };
const SPOTS: Spot[] = [
  "left-half",
  "right-half",
  "top-left",
  "top-right",
  "bottom-left",
  "bottom-right",
  "whole",
];

describe("resolveTarget", () => {
  it("lands on the picker's chosen slide, not the active one", async () => {
    const { presentation, helpers } = await bootPpt();
    helpers.selectSlide(presentation.slides[0]!.id);
    const target: InsertTarget = {
      slideId: presentation.slides[2]!.id,
      where: "free",
    };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    expect(resolved.slideId).toBe(presentation.slides[2]!.id);
  });

  it("falls back to the active slide, and its old error, with none selected", async () => {
    const { helpers } = await bootPpt();
    helpers.clearSelection();
    const target: InsertTarget = { slideId: null, where: "free" };
    await expect(
      PowerPoint.run((context) =>
        resolveTarget(context, "insert test", target, SIZE),
      ),
    ).rejects.toThrow("insert test: select a slide first.");
  });

  describe.each(SPOTS)("spot %s", (spot) => {
    it("lands inside its rectangle", async () => {
      const { presentation } = await bootPpt();
      const slideId = presentation.slides[0]!.id;
      // The same margin and gap placement.ts keeps: SLIDE_MARGIN 36,
      // SLIDE_GAP 12 (private there, so the check reuses layout.ts's own
      // spotBox instead of duplicating the numbers blind).
      const rect = spotBox(spot, SLIDE, 36, 12);
      const resolved = await PowerPoint.run((context) =>
        resolveTarget(context, "insert test", { slideId, where: spot }, SIZE),
      );
      const box = resolved.placement.box;
      expect(box.left).toBeGreaterThanOrEqual(rect.left - 0.01);
      expect(box.top).toBeGreaterThanOrEqual(rect.top - 0.01);
      expect(box.left + box.width).toBeLessThanOrEqual(
        rect.left + rect.width + 0.01,
      );
      expect(box.top + box.height).toBeLessThanOrEqual(
        rect.top + rect.height + 0.01,
      );
    });
  });

  it("flags overlapping when a spot already holds something", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      left: 36,
      top: 36,
      width: 400,
      height: 400,
    });
    const target: InsertTarget = { slideId: slide.id, where: "left-half" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    expect(resolved.placement.overlapping).toBe(true);
  });

  it("does not flag a spot nothing else occupies", async () => {
    const { presentation } = await bootPpt();
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      left: 36,
      top: 36,
      width: 400,
      height: 400,
    });
    const target: InsertTarget = { slideId: slide.id, where: "right-half" };
    const resolved = await PowerPoint.run((context) =>
      resolveTarget(context, "insert test", target, SIZE),
    );
    expect(resolved.placement.overlapping).toBe(false);
  });

  describe("selected-shape", () => {
    it("throws when nothing is selected", async () => {
      const { presentation } = await bootPpt();
      const target: InsertTarget = {
        slideId: presentation.slides[0]!.id,
        where: "selected-shape",
      };
      await expect(
        PowerPoint.run((context) =>
          resolveTarget(context, "insert test", target, SIZE),
        ),
      ).rejects.toThrow(
        "Select a shape on the slide first, or choose another spot.",
      );
    });

    it("fits into the selected shape's box and flags it for consumption when empty", async () => {
      const { presentation, helpers } = await bootPpt();
      const slide = presentation.slides[0]!;
      const placeholder = presentation.addShape(slide, {
        type: "Placeholder",
        hasText: false,
        left: 100,
        top: 80,
        width: 300,
        height: 150,
      });
      helpers.selectShapes([placeholder.id]);
      const target: InsertTarget = {
        slideId: slide.id,
        where: "selected-shape",
      };
      const resolved = await PowerPoint.run((context) =>
        resolveTarget(context, "insert test", target, {
          width: 600,
          height: 300,
        }),
      );
      expect(resolved.consume).toBe(placeholder.id);
      expect(resolved.placement.overlapping).toBe(false);
      expect(resolved.placement.box.width).toBeLessThanOrEqual(300);
      expect(resolved.placement.box.height).toBeLessThanOrEqual(150);
      // resolveTarget only flags it; nothing is deleted until finishTarget.
      expect(slide.shapes.map((one) => one.id)).toContain(placeholder.id);
    });

    it("keeps a non-empty selected shape and reports the overlap", async () => {
      const { presentation, helpers } = await bootPpt();
      const slide = presentation.slides[0]!;
      const picture = presentation.addShape(slide, {
        type: "Image",
        left: 50,
        top: 50,
        width: 200,
        height: 100,
      });
      helpers.selectShapes([picture.id]);
      const target: InsertTarget = {
        slideId: slide.id,
        where: "selected-shape",
      };
      const resolved = await PowerPoint.run((context) =>
        resolveTarget(context, "insert test", target, SIZE),
      );
      expect(resolved.consume).toBeUndefined();
      expect(resolved.placement.overlapping).toBe(true);
    });
  });
});

describe("finishTarget", () => {
  it("selects the slide the picker named, and deletes the consumed shape", async () => {
    const { presentation, helpers } = await bootPpt();
    const [first, , third] = presentation.slides;
    helpers.selectSlide(first!.id);
    const placeholder = presentation.addShape(third!, { type: "Placeholder" });
    const kept = presentation.addShape(third!, { type: "TextBox", text: "x" });

    await finishTarget(
      { slideId: third!.id, where: "free" },
      third!.id,
      placeholder.id,
    );

    expect(presentation.selectedSlideIds).toEqual([third!.id]);
    const remaining = third!.shapes.map((one) => one.id);
    expect(remaining).not.toContain(placeholder.id);
    expect(remaining).toContain(kept.id);
  });

  it("leaves 'This slide' alone and spends no round trip with nothing to do", async () => {
    const { presentation, helpers } = await bootPpt();
    helpers.selectSlide(presentation.slides[0]!.id);
    const before = helpers.syncCount();

    await finishTarget(
      { slideId: null, where: "free" },
      presentation.slides[0]!.id,
    );

    expect(helpers.syncCount()).toBe(before);
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[0]!.id]);
  });
});

// End to end through links.insertFromInbox, so each of the four payload
// kinds is proven to honour a non-default target, not only resolveTarget in
// isolation.
describe("insert paths honour the target", () => {
  it("the picture path lands on the chosen slide and selects it", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(400, 200));
    const target: InsertTarget = {
      slideId: presentation.slides[2]!.id,
      where: "free",
    };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    expect(placed.slideId).toBe(presentation.slides[2]!.id);
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[2]!.id]);
    expect(presentation.slides[2]!.shapes.map((one) => one.id)).toContain(
      placed.shapeId,
    );
  });

  it("the table path lands inside a chosen spot", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable([[{ t: "A" }, { t: "B" }]], [100, 100]);
    const target: InsertTarget = {
      slideId: presentation.slides[0]!.id,
      where: "top-right",
    };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    const table = presentation.slides[0]!.shapes.find(
      (one) => one.id === placed.shapeId,
    )!;
    expect(table.left).toBeGreaterThanOrEqual(486 - 0.01);
    expect(table.top).toBeGreaterThanOrEqual(36 - 0.01);
  });

  it("the text path consumes a selected empty placeholder and keeps its neighbour", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedText("Revenue: 1 234");
    const slide = presentation.slides[0]!;
    const placeholder = presentation.addShape(slide, {
      type: "Placeholder",
      hasText: false,
      left: 40,
      top: 40,
      width: 300,
      height: 60,
    });
    const kept: FakePptShape = presentation.addShape(slide, {
      type: "TextBox",
      text: "keep me",
      left: 10,
      top: 10,
      width: 40,
      height: 20,
    });
    helpers.selectShapes([placeholder.id]);
    const target: InsertTarget = { slideId: slide.id, where: "selected-shape" };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    const ids = slide.shapes.map((one) => one.id);
    expect(ids).not.toContain(placeholder.id);
    expect(ids).toContain(kept.id);
    expect(ids).toContain(placed.shapeId);
  });

  it("the chart path lands on the chosen slide", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(3), fakePng(400, 300));
    const target: InsertTarget = {
      slideId: presentation.slides[1]!.id,
      where: "whole",
    };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    expect(placed.slideId).toBe(presentation.slides[1]!.id);
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[1]!.id]);
  });
});
