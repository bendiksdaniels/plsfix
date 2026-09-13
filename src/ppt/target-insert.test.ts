// End to end through links.insertFromInbox: every insert path (picture -
// both with and without in-place refresh, table, text, chart, chart's
// draw-timeout fallback) honouring a non-default InsertTarget, and a failed
// insert never taking a consumed placeholder down with it. Split out of
// target.test.ts, which keeps resolveTarget/finishTarget in isolation.
// Strict load semantics are on throughout.

import { afterEach, describe, expect, it, vi } from "vitest";
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
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
} from "../../test/fakeppt";
import { fakePng } from "../../test/fakepng";
import { settleHungSync } from "../../test/hung-sync";
import { SLIDE_MARGIN, type InsertTarget } from "./placement";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
  vi.useRealTimers();
});

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
    expect(table.top).toBeGreaterThanOrEqual(SLIDE_MARGIN - 0.01);
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

  // Below PowerPointApi 1.8: the picture goes in through the Office selection
  // API (picture.ts), the other route insertLink takes - proving the target
  // reaches that branch too, not only the fill.setImage one above.
  it("the picture route below in-place-refresh support honours the target", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(200, 100));
    const target: InsertTarget = {
      slideId: presentation.slides[2]!.id,
      where: "top-left",
    };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    expect(placed.slideId).toBe(presentation.slides[2]!.id);
    expect(presentation.slides[2]!.shapes[0]!.type).toBe("Image");
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[2]!.id]);
  });

  // A failed insert must never take the placeholder down with it: resolveTarget
  // can confirm consume before the draw itself is the thing that fails.
  it("keeps the placeholder when the insert fails after resolveTarget confirms it", async () => {
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
    helpers.selectShapes([placeholder.id]);
    const target: InsertTarget = { slideId: slide.id, where: "selected-shape" };
    // resolveTarget's selected-shape path spends 2 syncs (the selection, then
    // the parent slide plus the placeholder check); the 3rd is the text
    // box's own confirming sync - fail exactly that one.
    helpers.failNextSync(new Error("PowerPoint stopped answering."), 2);

    await expect(
      links.insertFromInbox(item, ws, relay, target),
    ).rejects.toThrow();

    expect(slide.shapes.map((one) => one.id)).toContain(placeholder.id);
  });

  // The host swallows the chart draw mid-batch: the picture fallback still
  // has to honour a consume that resolveTarget already confirmed.
  it("the chart draw-timeout route still consumes a pending placeholder", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), fakePng(800, 400));
    const slide = presentation.slides[0]!;
    const placeholder = presentation.addShape(slide, {
      type: "Placeholder",
      hasText: false,
      left: 180,
      top: 120,
      width: 600,
      height: 300,
    });
    const kept = presentation.addShape(slide, {
      type: "TextBox",
      text: "keep me",
    });
    helpers.selectShapes([placeholder.id]);
    const target: InsertTarget = { slideId: slide.id, where: "selected-shape" };
    // Same 2 pre-draw syncs as above; the 3rd is the first SHAPES_PER_SYNC
    // chunk of the draw, which this leaves hanging.
    helpers.hangNextSync(3);
    vi.useFakeTimers();

    const placed = await settleHungSync(
      links.insertFromInbox(item, ws, relay, target),
    );

    const ids = slide.shapes.map((one) => one.id);
    expect(ids).not.toContain(placeholder.id);
    expect(ids).toContain(kept.id);
    expect(ids).toContain(placed.shapeId);
    expect(placed.note).toBe(
      "as a picture: PowerPoint stopped answering while drawing the shapes",
    );
  });
});
