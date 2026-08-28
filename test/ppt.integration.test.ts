// The PowerPoint side end to end against the fake host and the fake relay:
// inbox insert, scan by tags, status, in-place refresh, break, and the
// reinsertion fallback on hosts below PowerPointApi 1.8. Strict load semantics
// are on, so a scalar read the adapter forgot to load() fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type Payload,
} from "../src/link/model";
import { fitToSlide } from "../src/link/status";
import { createWorkspace, type KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "./fakeppt";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "SMT_LINK_00000000",
};

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

async function seedLink(
  png: string,
  workbook = src.workbook,
): Promise<InboxItem> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
  return {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

async function pushAgain(
  item: InboxItem,
  png: string,
  workbook = src.workbook,
): Promise<void> {
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 1,
    height: 1,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "1".repeat(64),
  };
  const keys = await deriveLinkKeys(item.token);
  await relay.putLink(
    item.id,
    keys.auth,
    await seal(keys.enc, item.id, encodePayload(payload)),
  );
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 3 });
  presentation = host.presentation;
  helpers = host.helpers;
  relay = new FakeRelay();
  links = await import("../src/ppt/links");
  helpers.selectSlide(presentation.slides[0]!.id);
});
afterEach(() => {
  uninstallFakePpt();
});

describe("insert from inbox", () => {
  it("creates a tagged picture-filled rectangle sized to the image and clears the inbox", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await relay.postInbox(
      ws.id,
      ws.auth,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );
    expect(await links.listInbox(ws, relay)).toHaveLength(1);
    await links.insertFromInbox(item, ws, relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect(shape.fillImage).toBe(fakePng(800, 400));
    expect(shape.lineVisible).toBe(false);
    expect({
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    }).toEqual(fitToSlide(800, 400));
    expect(JSON.parse(shape.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "range",
      rev: 1,
    });
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(await links.listInbox(ws, relay)).toHaveLength(0);
  });

  it("skips an inbox blob sealed with another workspace key, and an empty deck summarizes to nothing", async () => {
    const ws = await createWorkspace(memoryStore());
    const other = await createWorkspace(memoryStore());
    const mine = await seedLink(fakePng(10, 10));
    const theirs = await seedLink(fakePng(10, 10));
    await relay.postInbox(
      ws.id,
      ws.auth,
      mine.id,
      await seal(ws.enc, ws.id, encodeInboxItem(mine)),
    );
    await relay.postInbox(
      ws.id,
      ws.auth,
      theirs.id,
      await seal(other.enc, other.id, encodeInboxItem(theirs)),
    );
    const inbox = await links.listInbox(ws, relay);
    expect(inbox.map((entry) => entry.id)).toEqual([mine.id]);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(links.summarize(summary)).toBe("No links found");
  });
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
