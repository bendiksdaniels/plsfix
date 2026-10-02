// Hunt pass 1: src/link/local-store.ts's eviction: oldest-pasted first (not
// insertion order), links before inbox rows, eviction surviving a reopen, and a
// real deck's Update all over one kept and one evicted link.

import { afterEach, describe, expect, it } from "vitest";
import type { Bundle } from "../../src/link/bundle";
import { deriveLinkKeys, newToken, seal } from "../../src/link/crypto";
import { LOCAL_REV_BASE } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import {
  memoryPersistence,
  type LocalPersistence,
  type StoredLocalInbox,
  type StoredLocalLink,
} from "../../src/link/local-persist";
import { LocalStore } from "../../src/link/local-store";
import {
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
  type Payload,
} from "../../src/link/model";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "../fakeppt";
import { fakePng } from "../fakepng";
import * as links from "../../src/ppt/links";

const A = "a".repeat(32);
const B = "b".repeat(32);
const R = (n: number): number => LOCAL_REV_BASE + n;
const bytes = (n: number, fill = 1): Uint8Array => new Uint8Array(n).fill(fill);

afterEach(() => {
  uninstallFakePpt();
});

function bundle(
  links: [string, number, number][],
  inbox: string[] = [],
): Bundle {
  return {
    links: links.map(([id, rev, size]) => ({
      id,
      rev,
      sentAt: rev,
      blob: bytes(size, rev % 250),
    })),
    inbox: inbox.map((id) => ({ id, createdAt: 1, blob: bytes(4) })),
  };
}

describe("LocalStore.ingest: eviction under budget", () => {
  it("evicts the oldest of several old links, keeping the newest and the one in progress", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), {
      now: () => (now += 1),
      budget: 20,
    });
    const C = "c".repeat(32);
    await store.ingest(bundle([[A, R(1), 8]]), new Set());
    await store.ingest(bundle([[B, R(1), 8]]), new Set());
    // A third, unrelated paste that must evict the OLDEST (A) first, never
    // touching the link this very call is pasting (C).
    await store.ingest(bundle([[C, R(1), 8]]), new Set());

    const [a, b, c] = await store.status([
      { id: A, auth: "x" },
      { id: B, auth: "x" },
      { id: C, auth: "x" },
    ]);
    expect(a?.rev).toBeNull();
    expect(b?.rev).toBe(R(1));
    expect(c?.rev).toBe(R(1));
  });

  it("evicts stale inbox rows once the budget is spent on links, oldest first", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), {
      now: () => (now += 1),
      budget: 10,
    });
    const oldRow = "c".repeat(32);
    const newRow = "d".repeat(32);
    // Two inbox rows pasted a step apart, neither ever inserted on a deck.
    await store.ingest(bundle([], [oldRow]), new Set());
    await store.ingest(bundle([], [newRow]), new Set());
    // A small link's arrival tips the total over budget (4 + 4 + 4 = 12 > 10):
    // enough room is freed by dropping the OLDEST row alone (back to 8).
    await store.ingest(bundle([[A, R(1), 4]]), new Set());

    const inboxIds = (await store.listInbox("ws", "x")).map((row) => row.id);
    expect(inboxIds).not.toContain(oldRow);
    expect(inboxIds).toContain(newRow);
  });

  // The two tests above never distinguish a sort by pastedAt from a plain
  // walk in insertion order: A is pasted before B and evicted before B either
  // way, and the same is true of the two inbox rows. Both still pass with
  // both `.sort((a, b) => a.pastedAt - b.pastedAt)` calls deleted from
  // evict() (proven on this branch: 12/12 green with the sorts removed).
  // Re-pasting A after B moves A's pastedAt ahead of B's WITHOUT moving A in
  // the underlying Map's insertion order (Map.set on an existing key never
  // reorders it) - the one case where a real sort and a walk in insertion
  // order disagree about which link is actually oldest.
  it("re-pasting a link moves it to the back of the eviction queue, not just its own slot", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), {
      now: () => (now += 1),
      budget: 20,
    });
    const C = "c".repeat(32);
    await store.ingest(bundle([[A, R(1), 8]]), new Set()); // pastedAt 1
    await store.ingest(bundle([[B, R(1), 8]]), new Set()); // pastedAt 2
    // A again, unchanged rev: mergeRevision only refreshes pastedAt (to 3),
    // it does not touch the Map's key order.
    await store.ingest(bundle([[A, R(1), 8]]), new Set());
    // Tips the total to 24 > 20: by pastedAt, B (2) is now the oldest link
    // and is the one evicted, not A (whose insertion position is still
    // first). Sorting removed, the insertion-order walk evicts A instead.
    await store.ingest(bundle([[C, R(1), 8]]), new Set());

    const [a, b, c] = await store.status([
      { id: A, auth: "x" },
      { id: B, auth: "x" },
      { id: C, auth: "x" },
    ]);
    expect([a?.rev, b?.rev, c?.rev]).toEqual([R(1), null, R(1)]);
  });

  // The two eviction tests above also never give links AND inbox rows a
  // reason to compete for the same budget: evict() always finishes its walk
  // over links, checking the budget after each one, before it ever looks at
  // rows - so a row far older than a link is left alone whenever evicting
  // the link alone already buys back enough room. Swap that order (rows
  // before links) and the still-fits row would be the one dropped instead,
  // while the actually-over-budget link survives.
  it("evicts an old link before it ever touches an old inbox row that would have fit", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), {
      now: () => (now += 1),
      budget: 10,
    });
    const C = "c".repeat(32);
    const oldRow = "e".repeat(32);
    // A and the row arrive together, both stale by the time C lands.
    await store.ingest(bundle([[A, R(1), 4]], [oldRow]), new Set());
    // Total goes to 13 (4 + 4 + 5) > 10: evicting A alone (back to 9) is
    // enough, so a correct evict() never has to touch the row at all.
    await store.ingest(bundle([[C, R(1), 5]]), new Set());

    const [a] = await store.status([{ id: A, auth: "x" }]);
    expect(a?.rev).toBeNull();
    expect((await store.listInbox("ws", "x")).map((row) => row.id)).toEqual([
      oldRow,
    ]);
  });
});

describe("LocalStore.ingest: eviction survives a reopen", () => {
  // Keeps what it is given, so a second LocalStore.open reads it back - the
  // same shape src/link/local-store.test.ts's own reopen test uses, so a
  // link evict() drops only from the in-memory Map (never from `persist`
  // itself) would resurrect on this very reopen.
  function mapPersistence(): LocalPersistence {
    const linkRows = new Map<string, StoredLocalLink>();
    const inboxRows = new Map<string, StoredLocalInbox>();
    return {
      durable: true,
      load: async () => ({
        links: [...linkRows.values()],
        inbox: [...inboxRows.values()],
      }),
      putLinks: async (list) => list.forEach((l) => linkRows.set(l.id, l)),
      putInbox: async (list) => list.forEach((r) => inboxRows.set(r.id, r)),
      deleteLinks: async (ids) => ids.forEach((id) => linkRows.delete(id)),
      deleteInbox: async (ids) => ids.forEach((id) => inboxRows.delete(id)),
      clear: async () => {
        linkRows.clear();
        inboxRows.clear();
      },
    };
  }

  it("an evicted link does not come back once the same persistence reopens", async () => {
    const persist = mapPersistence();
    let now = 0;
    const first = await LocalStore.open(persist, {
      now: () => (now += 1),
      budget: 10,
    });
    await first.ingest(bundle([[A, R(1), 6]]), new Set());
    await first.ingest(bundle([[B, R(1), 6]]), new Set());
    const [beforeA, beforeB] = await first.status([
      { id: A, auth: "x" },
      { id: B, auth: "x" },
    ]);
    expect(beforeA?.rev).toBeNull();
    expect(beforeB?.rev).toBe(R(1));

    // A reopen calls persist.load(): if evict() had only cleared its own
    // Map and never called persist.deleteLinks, A would come back here.
    const reopened = await LocalStore.open(persist);
    const [afterA, afterB] = await reopened.status([
      { id: A, auth: "x" },
      { id: B, auth: "x" },
    ]);
    expect(afterA?.rev).toBeNull();
    expect(afterB?.rev).toBe(R(1));
  });
});

describe("LocalStore behind a real deck: Update all over one kept and one evicted link", () => {
  const PLACEHOLDER = fakePng(50, 25);
  const NEW_PICTURE = fakePng(400, 200);

  function newId(): string {
    return newLinkId((n) =>
      new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
    );
  }

  // Sealed the way Excel's LocalCollector actually records a copy, so
  // decodePayload inside updateLinks's own repaint has real bytes to open -
  // not the raw byte fills bundle()/ingest() are exercised with above.
  async function sealedBundleLink(
    collector: LocalCollector,
    workbook: string,
  ): Promise<{ id: string; token: string }> {
    const id = newId();
    const token = newToken();
    const keys = await deriveLinkKeys(token);
    const payload: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 400,
      height: 200,
      png: NEW_PICTURE,
      src: {
        workbook,
        sheet: "Model",
        ref: "B4:F12",
        anchor: "PLSFIX_LINK_00000000",
      },
      pushedAt: new Date().toISOString(),
      hash: "0".repeat(64),
    };
    await collector.putLink(
      id,
      keys.auth,
      await seal(keys.enc, id, encodePayload(payload)),
      1,
    );
    return { id, token };
  }

  // A minute in the past, comfortably behind sealedBundleLink's own
  // pushedAt (now): assertFresh (status.ts) refuses an "update" whose
  // payload is not newer than the tag already on the shape, and a plain
  // new Date().toISOString() here would race sealedBundleLink's on whichever
  // one lands in the same millisecond - flaky rather than wrong.
  const PLANTED_AT = new Date(Date.now() - 60_000).toISOString();

  // A shape already on the deck at rev 1 - one behind the local rev ingest()
  // is about to give its link - still showing the placeholder, never the
  // payload's own picture, so a later match against NEW_PICTURE actually
  // proves a repaint happened rather than the shape having started that way.
  function plantShape(
    presentation: FakePresentation,
    id: string,
    token: string,
    workbook: string,
    left: number,
  ): FakePptShape {
    const shape = presentation.addShape(presentation.slides[0]!, {
      type: "GeometricShape",
      fillImage: PLACEHOLDER,
      left,
      top: 0,
      width: 200,
      height: 100,
    });
    const tag: LinkTag = {
      v: 1,
      id,
      kind: "range",
      rev: 1,
      src: {
        workbook,
        sheet: "Model",
        ref: "B4:F12",
        anchor: "PLSFIX_LINK_00000000",
      },
      pushedAt: PLANTED_AT,
    };
    shape.tags.set(TAG_LINK, encodeTag(tag));
    shape.tags.set(TAG_KEY, token);
    return shape;
  }

  it("repaints the kept link and reports the evicted one as not pasted here, never a crash", async () => {
    const { presentation } = installFakePpt({ slides: 1 });

    // Gone first, kept second: a budget of 1 evicts anything not part of the
    // CURRENT ingest call the moment a second one arrives, so gone is
    // dropped as soon as kept is pasted, before Update all ever runs.
    const goneCollector = new LocalCollector();
    const gone = await sealedBundleLink(goneCollector, "Other.xlsx");
    const keptCollector = new LocalCollector();
    const kept = await sealedBundleLink(keptCollector, "Model_v5.xlsx");

    const store = await LocalStore.open(memoryPersistence(), { budget: 1 });
    await store.ingest(goneCollector.bundle(), new Set());
    await store.ingest(keptCollector.bundle(), new Set());

    const goneShape = plantShape(
      presentation,
      gone.id,
      gone.token,
      "Other.xlsx",
      0,
    );
    const keptShape = plantShape(
      presentation,
      kept.id,
      kept.token,
      "Model_v5.xlsx",
      250,
    );

    const rows = await links.listLinks(store);
    expect(rows).toHaveLength(2);

    const summary = await links.updateLinks(rows, store);

    expect(summary).toMatchObject({ updated: 1, notPasted: 1, failed: 0 });
    expect(keptShape.fillImage).toBe(NEW_PICTURE);
    expect(goneShape.fillImage).toBe(PLACEHOLDER);
  });
});
