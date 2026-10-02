// Scanning the deck for tagged links and reading their status against the
// fake relay: found through moves and copies, update/missing/wrong-key
// classification, and PowerPoint's active slide. Strict load semantics are
// on, so a scalar read the adapter forgot to load() fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { newToken } from "../src/link/crypto";
import { TAG_KEY } from "../src/link/model";
import type { InboxItem } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "./fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

// One full chunk of MAX_STATUS_ITEMS and a short second one: enough to prove
// both the split and that the tail is not silently dropped.
const MANY_LINKS = 250;

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

describe("scan and status", () => {
  it("finds a link after it was moved to another slide and after a copy", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const shapeId = presentation.slides[0]!.shapes[0]!.id;
    presentation.moveShape(shapeId, presentation.slides[2]!.id);
    presentation.copyShape(shapeId, presentation.slides[1]!.id);
    const rows = await links.listLinks(relay);
    expect(rows.map((r) => r.found.slideIndex).sort()).toEqual([1, 2]);
    expect(rows.every((r) => r.status === "current")).toBe(true);
  });

  it("reports update available, missing and wrong key", async () => {
    const ws = await createWorkspace(memoryStore());
    const a = await seedLink(fakePng(10, 10));
    const b = await seedLink(fakePng(10, 10));
    const c = await seedLink(fakePng(10, 10));
    for (const item of [a, b, c]) await links.insertFromInbox(item, ws, relay);
    await pushAgain(a, fakePng(10, 10));
    relay.links.delete(b.id);
    presentation.slides[0]!.shapes[2]!.tags.set(TAG_KEY, newToken());
    const rows = await links.listLinks(relay);
    expect(rows.map((r) => r.status)).toEqual([
      "updateAvailable",
      "missing",
      "wrongKey",
    ]);
  });

  // The invariant behind querying per (id, auth) rather than per id: one link
  // id can sit in the deck under two keys, and only one of them opens it.
  it("keeps a re-keyed copy apart from the original it shares an id with", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const original = presentation.slides[0]!.shapes[0]!;
    const copy = presentation.copyShape(
      original.id,
      presentation.slides[1]!.id,
    );
    copy.tags.set(TAG_KEY, newToken());
    await pushAgain(item, fakePng(800, 400));
    const rows = await links.listLinks(relay);
    expect(rows.map((r) => r.found.slideIndex)).toEqual([0, 1]);
    expect(rows.map((r) => r.status)).toEqual(["updateAvailable", "wrongKey"]);
    expect(rows[1]!.relayRev).toBeNull();
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, wrongKey: 1, failed: 0 });
    expect(original.setImageCalls).toBe(2);
    expect(copy.setImageCalls).toBe(1);
  });

  it("flags one unusable key as wrong and still reports its neighbours", async () => {
    const ws = await createWorkspace(memoryStore());
    const a = await seedLink(fakePng(10, 10));
    const b = await seedLink(fakePng(10, 10));
    for (const item of [a, b]) await links.insertFromInbox(item, ws, relay);
    presentation.slides[0]!.shapes[0]!.tags.set(TAG_KEY, "not a token");
    const rows = await links.listLinks(relay);
    expect(rows.map((r) => r.status)).toEqual(["wrongKey", "current"]);
    expect(rows[0]!.relayRev).toBeNull();
    expect(rows[1]!.relayRev).toBe(1);
  });

  // A pitch book past the relay's per-request ceiling. The server refuses a
  // longer batch with "400 too many items" (the fake does too), so before the
  // poll was chunked such a deck could not list a single row: listLinks threw
  // before the first one, and the Links tab, "Update all", "Revert last
  // update" and "Break link" were all dead with it.
  it("polls status in chunks of 200 and matches each answer to its own slice", async () => {
    const ws = await createWorkspace(memoryStore());
    const items: InboxItem[] = [];
    for (let index = 0; index < MANY_LINKS; index += 1) {
      const item = await seedLink(fakePng(10, 10));
      await links.insertFromInbox(item, ws, relay);
      items.push(item);
    }
    // The one row carrying an update sits in the second chunk, so an answer
    // read positionally against the whole list would land on the wrong link.
    const changed = MANY_LINKS - 7;
    await pushAgain(items[changed]!, fakePng(10, 10));
    const status = vi.spyOn(relay, "status");

    const rows = await links.listLinks(relay);

    expect(status.mock.calls.map(([batch]) => batch.length)).toEqual([200, 50]);
    expect(rows.map((row) => row.found.tag.id)).toEqual(
      items.map((item) => item.id),
    );
    expect(
      rows.flatMap((row, index) =>
        row.status === "updateAvailable" ? [index] : [],
      ),
    ).toEqual([changed]);
    expect(rows[changed]!.relayRev).toBe(2);
    expect(rows.filter((row) => row.status === "current")).toHaveLength(
      MANY_LINKS - 1,
    );
  });

  it("refuses a status answer that does not line up with the request", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    vi.spyOn(relay, "status").mockResolvedValue([]);
    await expect(links.listLinks(relay)).rejects.toThrow(
      "relay status: expected 1 rows, got 0",
    );
  });
});

describe("active slide", () => {
  it("reports PowerPoint's active slide, not a ticked link", async () => {
    const host = await import("../src/ppt/host");
    helpers.selectSlide(presentation.slides[1]!.id);
    expect(await host.activeSlideId()).toBe(presentation.slides[1]!.id);
  });

  it("reports null once the selection is cleared", async () => {
    const host = await import("../src/ppt/host");
    helpers.clearSelection();
    expect(await host.activeSlideId()).toBeNull();
  });
});
