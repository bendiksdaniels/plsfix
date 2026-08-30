// "Update all" on a deck the size of a real pitch book: 60 slides, 60 links,
// ten of them inside groups. What is asserted is the number of round trips to
// PowerPoint - the one cost the size of a deck must not multiply - and never
// the clock, which is measured and annotated for the record only. Strict load
// semantics are on, so a batch that reads a scalar it never loaded fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { InboxItem, Payload } from "../src/link/model";
import { FETCH_BLOB_CAP } from "../src/link/relay";
import { createWorkspace } from "../src/link/workspace";
import { REPAINT_BUDGET_BYTES } from "../src/ppt/batching";
import { SHAPES_PER_SYNC } from "../src/ppt/charts";
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
  pushAgain,
  pushChart,
  seedChart,
  seedLink,
  src,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const SLIDES = 60;
const GROUPED = 10;
const PNG = fakePng(800, 400);
// Padded to about 2.5 MiB sealed: two of these are past the response cap.
const BIG = fakePng(800, 400, 2_000_000);

// fakePng's header, in bytes: base64 spends four characters on every three.
const HEADER_BYTES = (fakePng(1, 1).length / 4) * 3;
// A full-slide render: exactly 5 MiB of base64, so three of them are three
// repaint batches and never one.
const HUGE_PNG_BYTES = 5 * 1024 * 1024;
const HUGE_PADDING = (HUGE_PNG_BYTES / 4) * 3 - HEADER_BYTES;

// One batched load for the slides, one for every slide's shapes, one per level
// of grouping and one for every shape's tags: four, whatever N is.
const SCAN_SYNC_BUDGET = 4;
// Sixty small pictures are far inside the repaint budget, so every in-place
// repaint of this update-all travels in one batch and one sync.
const UPDATE_SYNC_BUDGET = 1;
// A chart is drawn shape by shape, so its insert is bounded by the shapes it
// spends, not by the deck: the selected slide, the boxes already on it, one
// sync per SHAPES_PER_SYNC of the layout, and one that groups and tags them.
const CHART_POINTS = 6;
const CHART_SHAPES = 20;
const CHART_SYNC_BUDGET = 2 + Math.ceil(CHART_SHAPES / SHAPES_PER_SYNC) + 1; // 5

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

// A deck built the way a user builds one: a link inserted on each slide, then
// the first ten dragged into a group with a caption beside them.
async function seedDeck(): Promise<InboxItem[]> {
  const ws = await createWorkspace(memoryStore());
  while (presentation.slides.length < SLIDES) presentation.addSlide();
  const items: InboxItem[] = [];
  for (const slide of presentation.slides) {
    helpers.selectSlide(slide.id);
    const item = await seedLink(PNG);
    await links.insertFromInbox(item, ws, relay);
    items.push(item);
  }
  for (const slide of presentation.slides.slice(0, GROUPED)) {
    const picture = slide.shapes[0]!;
    const caption = presentation.addShape(slide, {
      left: 5,
      top: 5,
      width: 60,
    });
    presentation.groupShapes([picture.id, caption.id], slide.id);
  }
  return items;
}

function trips(syncs: number, ms: number): string {
  return `${String(syncs)} ${syncs === 1 ? "sync" : "syncs"} in ${ms.toFixed(0)} ms`;
}

// A repaint payload built here instead of pushed through the relay: what is
// under test is how many host batches the bytes buy, not another round of
// sealing and opening.
function hugePayload(png: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src,
    pushedAt: new Date().toISOString(),
    hash: "2".repeat(64),
  };
}

// The pictures themselves, group or no group: one per slide, in slide order.
function pictures(): FakePptShape[] {
  return presentation.slides.map((slide) => {
    const shape = slide.shapes[0]!;
    return shape.group ? shape.group.shapes[0]! : shape;
  });
}

describe("update all on a 60-slide deck", () => {
  it("scans and updates in a constant number of round trips", async ({
    annotate,
  }) => {
    const items = await seedDeck();
    for (const item of items) await pushAgain(item, PNG);
    const status = vi.spyOn(relay, "status");
    const fetchLinks = vi.spyOn(relay, "fetchLinks");
    const getLink = vi.spyOn(relay, "getLink");

    // Every sync listLinks costs is scanLinks': the relay round trip after it
    // never touches PowerPoint.
    const scanFrom = helpers.syncCount();
    const scanStart = performance.now();
    const rows = await links.listLinks(relay);
    const scanSyncs = helpers.syncCount() - scanFrom;
    const scanMs = performance.now() - scanStart;

    const updateFrom = helpers.syncCount();
    const updateStart = performance.now();
    const summary = await links.updateLinks(rows, relay);
    const updateSyncs = helpers.syncCount() - updateFrom;
    const updateMs = performance.now() - updateStart;

    await annotate(
      `${String(SLIDES)} slides, ${String(rows.length)} links (${String(GROUPED)} grouped): ` +
        `scan ${trips(scanSyncs, scanMs)}, update ${trips(updateSyncs, updateMs)}`,
    );

    expect(rows).toHaveLength(SLIDES);
    expect(summary).toMatchObject({ updated: SLIDES, failed: 0 });
    expect(scanSyncs).toBeLessThanOrEqual(SCAN_SYNC_BUDGET);
    expect(updateSyncs).toBeLessThanOrEqual(UPDATE_SYNC_BUDGET);
    // The network is the rest of the cost, and the size of the deck must not
    // multiply it either: one status poll and one batched fetch, whatever N
    // is. The syncs above are on top of these, not instead of them.
    expect(status).toHaveBeenCalledTimes(1);
    expect(fetchLinks).toHaveBeenCalledTimes(1);
    expect(getLink).not.toHaveBeenCalled();
    expect(pictures().map((shape) => shape.setImageCalls)).toEqual(
      new Array<number>(SLIDES).fill(2),
    );
  });

  // Blobs the size of a real full-slide render: two of them are past the
  // relay's response cap, so the second comes back deferred rather than
  // truncated, and the per-row GET the batch replaced fetches it.
  it("defers what does not fit in one response and fetches it on its own", async ({
    annotate,
  }) => {
    const ws = await createWorkspace(memoryStore());
    const items: InboxItem[] = [];
    for (const slide of presentation.slides.slice(0, 2)) {
      helpers.selectSlide(slide.id);
      const item = await seedLink(PNG);
      await links.insertFromInbox(item, ws, relay);
      items.push(item);
    }
    for (const item of items) await pushAgain(item, BIG);
    const stored = items.map((item) => relay.links.get(item.id)!.blob.length);
    const total = stored.reduce((sum, size) => sum + size, 0);
    expect(total).toBeGreaterThan(FETCH_BLOB_CAP);
    await annotate(`2 pictures, ${(total / 1024 / 1024).toFixed(1)} MiB total`);

    const rows = await links.listLinks(relay);
    const fetchLinks = vi.spyOn(relay, "fetchLinks");
    const getLink = vi.spyOn(relay, "getLink");
    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect(fetchLinks).toHaveBeenCalledTimes(1);
    expect(getLink).toHaveBeenCalledTimes(1);
    expect(
      presentation.slides
        .slice(0, 2)
        .map((slide) => slide.shapes[0]!.setImageCalls),
    ).toEqual([2, 2]);
  });

  // What bounds a repaint batch is bytes, not rows: three full-slide pictures
  // are past the budget together, so they travel in three host requests
  // instead of holding 15 MiB in one.
  it("cuts the repaint into one host batch per 8 MiB of payload", async ({
    annotate,
  }) => {
    const ws = await createWorkspace(memoryStore());
    for (const slide of presentation.slides) {
      helpers.selectSlide(slide.id);
      await links.insertFromInbox(await seedLink(PNG), ws, relay);
    }
    const rows = await links.listLinks(relay);
    const host = await import("../src/ppt/host");
    const png = fakePng(800, 400, HUGE_PADDING);
    expect(png.length).toBe(HUGE_PNG_BYTES);
    expect(rows.length * png.length).toBeGreaterThan(REPAINT_BUDGET_BYTES);
    const batch = rows.map((row) => ({
      found: row.found,
      payload: hugePayload(png),
      rev: 2,
    }));

    const failures: unknown[] = [];
    const from = helpers.syncCount();
    const start = performance.now();
    const painted = await links.applyBatch(batch, host, (_found, error) => {
      failures.push(error);
    });
    const syncs = helpers.syncCount() - from;
    await annotate(
      `${String(batch.length)} payloads of ` +
        `${(png.length / 1024 / 1024).toFixed(1)} MiB: ` +
        trips(syncs, performance.now() - start),
    );

    expect(failures).toEqual([]);
    expect(painted).toBe(rows.length);
    expect(syncs).toBe(rows.length);
    expect(pictures().map((shape) => shape.setImageCalls)).toEqual(
      new Array<number>(rows.length).fill(2),
    );
  });

  // A chart is the one link made of many shapes, and the web charges per
  // shape and per round trip: what is bounded here is the syncs the drawing
  // costs, chunked at SHAPES_PER_SYNC, never one per shape.
  it("draws and redraws a chart in one sync per chunk plus one", async ({
    annotate,
  }) => {
    const ws = await createWorkspace(memoryStore());
    const chart = columnChart(CHART_POINTS);
    const item = await seedChart(chart, PNG);

    const insertFrom = helpers.syncCount();
    const insertStart = performance.now();
    await links.insertFromInbox(item, ws, relay);
    const insertSyncs = helpers.syncCount() - insertFrom;
    const insertMs = performance.now() - insertStart;

    const group = presentation.slides[0]!.shapes[0]!;
    expect(group.type).toBe("Group");
    expect(group.group!.shapes).toHaveLength(CHART_SHAPES);

    await pushChart(item, chart, PNG);
    const rows = await links.listLinks(relay);
    const repaintFrom = helpers.syncCount();
    const repaintStart = performance.now();
    const summary = await links.updateLinks(rows, relay);
    const repaintSyncs = helpers.syncCount() - repaintFrom;

    await annotate(
      `${String(CHART_SHAPES)} shapes: insert ${trips(insertSyncs, insertMs)}, ` +
        `repaint ${trips(repaintSyncs, performance.now() - repaintStart)}`,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(insertSyncs).toBe(CHART_SYNC_BUDGET);
    // A repaint knows the box already, so it spends the drawing syncs alone.
    expect(repaintSyncs).toBe(CHART_SYNC_BUDGET - 2);
  });

  // The batch is a speed-up, not a new failure mode: a shape the host refuses
  // costs its own row and nothing else.
  it("keeps one refused shape from blocking the rest of the batch", async () => {
    const items = await seedDeck();
    for (const item of items) await pushAgain(item, PNG);
    const rows = await links.listLinks(relay);
    // The shape is gone from the deck, so its refresh - batched or not - throws.
    presentation.deleteShape(rows[3]!.found.shapeId);

    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: SLIDES - 1, failed: 1 });
    expect(summary.failures).toHaveLength(1);
  });
});
