// The Mac gate in charts.ts: PowerPoint for Mac never gets a shape chart, on
// insert or on refresh, and the pane says why beside the picture it got.
import { afterEach, describe, expect, it } from "vitest";
import { createWorkspace } from "../link/workspace";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
} from "../../test/ppt.support";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { fakePng } from "../../test/fakepng";
import { CHARTS_MAC_PICTURE, hostDrawsCharts } from "./charts";
import { DEFAULT_TARGET } from "./placement";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

describe("PowerPoint for Mac keeps every chart a picture", () => {
  it("draws on the Windows desktop and on the web, never on the Mac", async () => {
    const { helpers } = await bootPpt();
    expect(hostDrawsCharts()).toBe(true);
    helpers.setPlatform("OfficeOnline");
    expect(hostDrawsCharts()).toBe(true);
    helpers.setPlatform("Mac");
    expect(hostDrawsCharts()).toBe(false);
  });

  it("inserts the picture on a Mac, with the reason beside it", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    helpers.setPlatform("Mac");
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(3), fakePng(400, 300));

    const placed = await links.insertFromInbox(item, ws, relay, DEFAULT_TARGET);

    expect(placed.note).toBe(CHARTS_MAC_PICTURE);
    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(1);
    expect(shapes[0]!.type).not.toBe("Group");
    expect(shapes[0]!.fillImage).toBeTruthy();
  });

  it("the same item is a shape group on the Windows desktop", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(3), fakePng(400, 300));

    const placed = await links.insertFromInbox(item, ws, relay, DEFAULT_TARGET);

    expect(placed.note).toBeUndefined();
    expect(presentation.findShape(placed.shapeId).shape.type).toBe("Group");
  });
});
