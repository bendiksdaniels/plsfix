// The Mac tier in chart-draw.ts: PowerPoint for Mac takes a chart's group as
// sub-groups of GROUP_TIER_MAC members under the link's group, every other
// host as the one flat group, and a failed top group leaves nothing behind.
import { afterEach, describe, expect, it } from "vitest";
import { createWorkspace } from "../link/workspace";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
} from "../../test/ppt.support";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
} from "../../test/fakeppt";
import { fakePng } from "../../test/fakepng";
import { GROUP_TIER_MAC, groupTier } from "./chart-draw";
import { DEFAULT_TARGET } from "./placement";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

// One title, six bars, six value labels, a baseline and six category labels.
const COLUMN_SHAPES = 20;
// The insert's syncs on the Mac: select the slide, place, two chunks of
// twelve and eight, the sub-groups, then the link's own group.
const TOP_GROUP_SYNC = 5;

function leaves(shape: FakePptShape): FakePptShape[] {
  return shape.type === "Group" ? shape.group!.shapes.flatMap(leaves) : [shape];
}

function fanout(shape: FakePptShape): number {
  if (shape.type !== "Group" || shape.group === null) return 0;
  return Math.max(shape.group.shapes.length, ...shape.group.shapes.map(fanout));
}

async function insertOn(platform: string) {
  const booted = await bootPpt();
  booted.helpers.setPlatform(platform);
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(columnChart(6), fakePng(800, 400));
  const placed = await booted.links.insertFromInbox(
    item,
    ws,
    booted.relay,
    DEFAULT_TARGET,
  );
  return { ...booted, item, placed };
}

describe("PowerPoint for Mac groups a chart in tiers", () => {
  it("tiers on the Mac only", async () => {
    const { helpers } = await bootPpt();
    expect(groupTier()).toBeNull();
    helpers.setPlatform("OfficeOnline");
    expect(groupTier()).toBeNull();
    helpers.setPlatform("Mac");
    expect(groupTier()).toBe(GROUP_TIER_MAC);
  });

  it("the Mac link is a group of sub-groups of six with every shape inside", async () => {
    const { presentation, placed } = await insertOn("Mac");
    expect(placed.note).toBeUndefined();
    const slide = presentation.slides[0]!;
    expect(slide.shapes.map((shape) => shape.id)).toEqual([placed.shapeId]);
    const link = slide.shapes[0]!;
    expect(link.type).toBe("Group");
    const subs = link.group!.shapes;
    expect(subs.map((sub) => sub.type)).toEqual([
      "Group",
      "Group",
      "Group",
      "Group",
    ]);
    expect(subs.map((sub) => sub.group!.shapes.length)).toEqual([6, 6, 6, 2]);
    expect(leaves(link)).toHaveLength(COLUMN_SHAPES);
  });

  // Twelve columns are 38 shapes: one addGroup of seven sub-groups of six,
  // which is still under the 19 that killed 16.107, but a 40-point chart is
  // 21 sub-groups and that is not. Recurse until no group has more than six.
  it("never addGroups more than six members, even on a wide chart", async () => {
    const booted = await bootPpt();
    booted.helpers.setPlatform("Mac");
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(12), fakePng(800, 400));
    const placed = await booted.links.insertFromInbox(
      item,
      ws,
      booted.relay,
      DEFAULT_TARGET,
    );
    expect(placed.note).toBeUndefined();
    const link = booted.presentation.findShape(placed.shapeId).shape;
    expect(fanout(link)).toBeLessThanOrEqual(GROUP_TIER_MAC);
    expect(leaves(link).length).toBeGreaterThan(COLUMN_SHAPES);
  });

  it("the Windows desktop keeps the one flat group", async () => {
    const { presentation, placed } = await insertOn("PC");
    const link = presentation.findShape(placed.shapeId).shape;
    expect(link.group!.shapes).toHaveLength(COLUMN_SHAPES);
    expect(link.group!.shapes.some((one) => one.type === "Group")).toBe(false);
  });

  it("Update all on the Mac redraws in tiers and keeps one link", async () => {
    const host = await import("./host");
    const { links, presentation, relay, item, placed } = await insertOn("Mac");
    await pushChart(item, columnChart(7), fakePng(800, 400));
    const rows = await links.listLinks(relay);

    const summary = await links.updateLinks(rows, relay, host);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const slide = presentation.slides[0]!;
    expect(slide.shapes).toHaveLength(1);
    const drawn = slide.shapes[0]!;
    expect(drawn.id).not.toBe(placed.shapeId);
    expect(drawn.group!.shapes.every((sub) => sub.type === "Group")).toBe(true);
    // Seven columns: a title, seven bars, seven values, a baseline, seven labels.
    expect(leaves(drawn)).toHaveLength(23);
  });

  it("a top group the host refuses takes its sub-groups down with it", async () => {
    const booted = await bootPpt();
    booted.helpers.setPlatform("Mac");
    booted.helpers.failNextSync(new Error("the host hung"), TOP_GROUP_SYNC);
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), fakePng(800, 400));

    await expect(
      booted.links.insertFromInbox(item, ws, booted.relay, DEFAULT_TARGET),
    ).rejects.toThrow("the host hung");

    expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
  });
});
