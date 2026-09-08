// A chart link whose Excel side could not describe the chart, or whose group
// this host can no longer draw: the picture lands or stays, and the pane says
// why - on insert as the note beside "Inserted", on update as a line under
// the counts. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
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
const ISSUE = "7 series; shapes draw up to 6";
const LABEL = "Model: B4:F12";

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
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

describe("insert a chart Excel could not describe", () => {
  it("inserts the picture and repeats Excel's reason", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(null, PNG, ISSUE);
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(shapes()[0]!.fillImage).toBe(PNG);
    expect(placed.note).toBe(`as a picture: ${ISSUE}`);
  });

  it("says nothing about a picture that carries no reason", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(null, PNG);
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toBeUndefined();
  });
});

describe("update a chart link that is now a picture", () => {
  it("names the link and Excel's reason under the counts", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    expect(shapes()[0]!.type).toBe("Group");
    await pushChart(item, null, PNG, ISSUE);

    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(summary.notes).toEqual([`${LABEL} as a picture: ${ISSUE}`]);
  });

  it("names this host's own reason when the chart outgrew its budget", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    helpers.setPlatform("OfficeOnline");
    await pushChart(item, columnChart(12), PNG);

    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(summary.notes).toEqual([
      `${LABEL} as a picture: 38 shapes is over this host's budget of 30`,
    ]);
  });

  it("stays quiet while the chart still draws", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    await pushChart(item, columnChart(7), PNG);

    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([]);
  });
});
