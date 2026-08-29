// Refreshing and breaking a tracked link against the fake host and the fake
// relay: in-place repaint, geometry left alone bar an aspect-ratio height
// change, a broken link's tags removed, and the reinsertion fallback on
// hosts below PowerPointApi 1.8. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { RelayError } from "../src/link/relay";
import { fitToSlide } from "../src/link/status";
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

describe("update", () => {
  it("repaints in place without touching geometry, bumps the tag rev", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    shape.left = 123;
    shape.top = 45; // the user moved it
    await pushAgain(item, fakePng(1600, 800));
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({
      updated: 1,
      current: 0,
      missing: 0,
      wrongKey: 0,
      failed: 0,
    });
    expect(shape.setImageCalls).toBe(2);
    expect([shape.left, shape.top, shape.width]).toEqual([
      123,
      45,
      fitToSlide(800, 400).width,
    ]);
    expect(JSON.parse(shape.tags.get(TAG_LINK)!).rev).toBe(2);
  });

  it("recomputes only the height when the aspect ratio changed", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    await pushAgain(item, fakePng(800, 800));
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(shape.width).toBe(fitToSlide(800, 400).width);
    expect(shape.height).toBe(shape.width);
  });

  it("flags a changed source workbook and counts up-to-date rows", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(10, 10), "Model_v5.xlsx");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary.sourceChanges).toEqual(["Model_v4.xlsx -> Model_v5.xlsx"]);
    const again = await links.updateLinks(await links.listLinks(relay), relay);
    expect(again.current).toBe(1);
    expect(links.summarize(summary)).toBe("1 updated");
  });

  // A relay that refuses the whole batch has said nothing about any one link,
  // and an older one does not know the route at all: every row falls back to
  // the GET it would have made, and the deck still updates.
  it("falls back to a GET per row when the relay refuses the batch", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(10, 10));
    const rows = await links.listLinks(relay);
    vi.spyOn(relay, "fetchLinks").mockRejectedValue(
      new RelayError("missing", "no such route", 404),
    );
    const getLink = vi.spyOn(relay, "getLink");

    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, missing: 0, failed: 0 });
    expect(getLink).toHaveBeenCalledTimes(1);
  });

  // The poll and the fetch are two round trips, and the relay can move
  // between them. Whatever the batch leaves out lands on the counter the row
  // would have had if the poll had seen it.
  it("maps what the batch left out onto the same counters", async () => {
    const ws = await createWorkspace(memoryStore());
    const seeded = [];
    for (const slide of presentation.slides) {
      helpers.selectSlide(slide.id);
      const item = await seedLink(fakePng(10, 10));
      await links.insertFromInbox(item, ws, relay);
      seeded.push(item);
    }
    for (const item of seeded) await pushAgain(item, fakePng(10, 10));
    const rows = await links.listLinks(relay);
    expect(rows.map((row) => row.status)).toEqual(
      new Array<string>(3).fill("updateAvailable"),
    );
    // One link is gone, one changed hands, one went back to the revision the
    // deck already holds.
    relay.links.delete(seeded[0]!.id);
    relay.links.get(seeded[1]!.id)!.auth = "another key";
    relay.links.get(seeded[2]!.id)!.rev = 1;

    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({
      updated: 0,
      current: 1,
      missing: 1,
      wrongKey: 1,
      failed: 0,
    });
    expect(links.summarize(summary)).toBe(
      "1 up to date, 1 missing, 1 wrong key",
    );
  });

  it("keeps the reason a row failed, and the other rows still update", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(10, 10));
    const refusing: LinksModule.PptHost = {
      scanLinks: () => Promise.resolve([]),
      insertLink: () => Promise.reject(new Error("unused")),
      refreshLink: () =>
        Promise.reject(new Error("PowerPoint refused the picture")),
      breakLink: () => Promise.resolve(),
      goToSlide: () => Promise.resolve(),
    };
    const rows = await links.listLinks(relay);
    const summary = await links.updateLinks(rows, relay, refusing);
    expect(summary.failed).toBe(1);
    expect(summary.failures).toEqual([
      "Model!B4:F12: PowerPoint refused the picture",
    ]);
    expect(links.summarize(summary)).toBe("1 failed");
  });
});

describe("break link", () => {
  it("removes the tags and keeps the picture", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    const host = await import("../src/ppt/host");
    await host.breakLink(row!.found);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect(shape.tags.size).toBe(0);
    expect(shape.fillImage).toBe(fakePng(10, 10));
    expect(await links.listLinks(relay)).toHaveLength(0);
  });
});

describe("hosts below PowerPointApi 1.8", () => {
  beforeEach(() => {
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
  });

  it("inserts through the selection and refreshes by reinsertion at the same box", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    expect(helpers.insertedViaSelection()).toHaveLength(1);
    const before = presentation.slides[0]!.shapes[0]!;
    expect(before.type).toBe("Image");
    expect(before.tags.get(TAG_KEY)).toBe(item.token);
    before.left = 50;
    await pushAgain(item, fakePng(800, 400));
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    const after = presentation.slides[0]!.shapes[0]!;
    expect(after.id).not.toBe(before.id);
    expect([after.left, after.top, after.width, after.height]).toEqual([
      50,
      before.top,
      before.width,
      before.height,
    ]);
    expect(JSON.parse(after.tags.get(TAG_LINK)!).rev).toBe(2);
  });

  // The reinsertion inserts before it deletes, so a host that refuses the new
  // picture cannot take the old one - and both tags - down with it.
  it("keeps the old picture and its tags when the reinsertion fails", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const before = presentation.slides[0]!.shapes[0]!;
    await pushAgain(item, fakePng(800, 400));
    helpers.failNextSelectionInsert("the host is out of memory");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([
      "refresh Model!B4:F12: the host is out of memory",
    ]);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(presentation.slides[0]!.shapes[0]).toBe(before);
    expect(before.fillImage).toBe(fakePng(800, 400));
    expect(before.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(before.tags.get(TAG_LINK)!).rev).toBe(1);
  });
});
