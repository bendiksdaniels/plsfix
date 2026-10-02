// A chart link whose chart data arrives unreadable - a schema the pane does
// not know, a field Excel wrote wrong - still inserts: the picture that
// travelled beside it lands and the note says why there are no shapes. Before
// src/link/chart-guard.ts the decoder refused the whole payload and the deck
// could not insert the link at all. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { CHART_UNREADABLE } from "../src/link/chart-guard";
import { createWorkspace } from "../src/link/workspace";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// The same size, a byte longer: what the next push of the same source sends.
const NEXT_PNG = fakePng(800, 400, 1);
const LABEL = "Model: B4:F12";
const NOTE = `as a picture: ${CHART_UNREADABLE}`;

// A chart object the validator refuses: the kind is one no slide draws, and
// the categories are not even a list. Cast because nothing typed can build
// it - this is what a newer Excel, or a corrupt blob, puts on the relay.
const BROKEN = {
  v: 1,
  kind: "radar",
  title: "Revenue",
  categories: "2024",
  series: [],
} as unknown as ChartData;

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

async function updateAll(): Promise<LinksModule.UpdateSummary> {
  return links.updateLinks(await links.listLinks(relay), relay);
}

describe("insert a chart link whose chart data is unreadable", () => {
  it("inserts the picture, tags it, and says the data was unreadable", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(BROKEN, PNG);
    const placed = await links.insertFromInbox(item, ws, relay);

    expect(shapes()).toHaveLength(1);
    const shape = shapes()[0]!;
    expect(shape.type).not.toBe("Group");
    expect(shape.fillImage).toBe(PNG);
    expect(placed.note).toBe(NOTE);
    // A picture with both tags is a link like any other: it lists, updates
    // and reverts, and the next good push draws the shapes.
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(shape.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "chart",
    });
  });

  it("lets the modeller's own reason stand when the payload carries one", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(BROKEN, PNG, "7 series; shapes draw up to 6");
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toBe("as a picture: 7 series; shapes draw up to 6");
  });

  it("repaints it as the picture it is when a later push reads fine", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(BROKEN, PNG);
    await links.insertFromInbox(item, ws, relay);
    expect(shapes()[0]!.type).not.toBe("Group");

    await pushChart(item, columnChart(6), NEXT_PNG);
    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    // A link inserted as a picture stays a picture: a repaint writes the
    // image, never the geometry, and only a group is ever redrawn as shapes.
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(shapes()[0]!.fillImage).toBe(NEXT_PNG);
    expect(summary.notes).toEqual([]);
  });

  it("keeps the link when a later push breaks the chart data", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    expect(shapes()[0]!.type).toBe("Group");

    await pushChart(item, BROKEN, PNG);
    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(summary.notes).toEqual([`${LABEL} ${NOTE}`]);
  });
});
