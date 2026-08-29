// "Update all" on a deck the size of a real pitch book: 60 slides, 60 links,
// ten of them inside groups. What is asserted is the number of round trips to
// PowerPoint - the one cost the size of a deck must not multiply - and never
// the clock, which is measured and annotated for the record only. Strict load
// semantics are on, so a batch that reads a scalar it never loaded fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../src/link/model";
import { FETCH_BLOB_CAP } from "../src/link/relay";
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
import { bootPpt, memoryStore, pushAgain, seedLink } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const SLIDES = 60;
const GROUPED = 10;
const PNG = fakePng(800, 400);
// Padded to about 2.5 MiB sealed: two of these are past the response cap.
const BIG = fakePng(800, 400, 2_000_000);

// One batched load for the slides, one for every slide's shapes, one per level
// of grouping and one for every shape's tags: four, whatever N is.
const SCAN_SYNC_BUDGET = 4;
// Every in-place repaint of an update-all travels in one batch.
const UPDATE_SYNC_BUDGET = 3;

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
