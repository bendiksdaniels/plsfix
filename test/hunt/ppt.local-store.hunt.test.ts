// Hunt pass 1: src/link/local-store.ts's merge rules and API branches: rollback
// protection two revisions deep, the same bundle pasted twice, getLink / getLinkRev
// on the current revision, inbox order, the writer refusals, fetchLinks, deleteInbox.

import { afterEach, describe, expect, it } from "vitest";
import type { Bundle } from "../../src/link/bundle";
import { LOCAL_REV_BASE } from "../../src/link/local";
import { memoryPersistence } from "../../src/link/local-persist";
import { LocalStore, NOT_PASTED_HERE } from "../../src/link/local-store";
import { uninstallFakePpt } from "../fakeppt";

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

describe("mergeRevision: rollback protection goes two deep", () => {
  it("an id pasted at rev 3 then 2 keeps 2 as 'previous'; pasting rev 1 next changes neither", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(3), 2]]), new Set());
    await store.ingest(bundle([[A, R(2), 2]]), new Set());
    // Sanity: rev 2 is now the previous slot (getLinkRev(2) must answer it).
    await expect(store.getLinkRev(A, "x", R(2))).resolves.toMatchObject({
      rev: R(2),
    });

    // A bundle older than BOTH the current (3) and the previous (2) slot.
    await store.ingest(bundle([[A, R(1), 2]]), new Set());

    const [status] = await store.status([{ id: A, auth: "x" }]);
    expect(status?.rev).toBe(R(3));
    // The previous slot must still be rev 2, never downgraded to rev 1.
    await expect(store.getLinkRev(A, "x", R(2))).resolves.toMatchObject({
      rev: R(2),
    });
    await expect(store.getLinkRev(A, "x", R(1))).rejects.toThrow(
      NOT_PASTED_HERE,
    );
  });
});

describe("mergeRevision: the exact-same bundle pasted twice", () => {
  it("keeps the one payload and only refreshes pastedAt, never re-adding a 'previous'", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());

    await expect(store.getLink(A, "x", R(1))).resolves.toBe("unchanged");
    // Still nothing at rev 0: a same-rev re-paste never manufactures a
    // "previous" slot out of thin air.
    await expect(store.getLinkRev(A, "x", R(0))).rejects.toThrow(
      NOT_PASTED_HERE,
    );
  });
});

describe("LocalStore.getLink / getLinkRev: the current-revision branch", () => {
  it("getLink answers 'unchanged' when the caller already has the current rev", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());

    await expect(store.getLink(A, "x", R(1))).resolves.toBe("unchanged");
    await expect(store.getLink(A, "x", R(99))).resolves.toMatchObject({
      rev: R(1),
    });
  });

  it("getLinkRev answers the CURRENT revision, not only the previous one", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());
    await store.ingest(bundle([[A, R(2), 2]]), new Set());

    await expect(store.getLinkRev(A, "x", R(2))).resolves.toEqual({
      rev: R(2),
      blob: bytes(2, R(2) % 250),
    });
  });
});

describe("LocalStore.listInbox: ordering with more than one row waiting", () => {
  it("sorts newest first by createdAt, not insertion order", async () => {
    const store = await LocalStore.open(memoryPersistence());
    const early = "c".repeat(32);
    const late = "d".repeat(32);
    await store.ingest(
      {
        links: [],
        inbox: [
          { id: early, createdAt: 100, blob: bytes(2) },
          { id: late, createdAt: 200, blob: bytes(2) },
        ],
      },
      new Set(),
    );

    const ids = (await store.listInbox("ws", "x")).map((row) => row.id);
    expect(ids).toEqual([late, early]);
  });
});

describe("LocalStore: the writer methods are Excel's, never this deck's", () => {
  it("putLink, postInbox, deleteLink and touchLinks all refuse with the same sentence", async () => {
    const store = await LocalStore.open(memoryPersistence());
    const refuses = "local mode writes through ingest()";

    await expect(store.putLink(A, "x", bytes(1))).rejects.toThrow(refuses);
    await expect(store.postInbox("ws", "x", A, bytes(1))).rejects.toThrow(
      refuses,
    );
    await expect(store.deleteLink(A, "x")).rejects.toThrow(refuses);
    await expect(store.touchLinks([{ id: A, auth: "x" }])).rejects.toThrow(
      refuses,
    );
  });
});

describe("LocalStore.fetchLinks: the three outcomes in one batch", () => {
  it("omits an unknown id as notPasted, an equal rev as unchanged, and hands back the rest", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());

    const result = await store.fetchLinks([
      { id: A, auth: "x", knownRev: R(1) },
      { id: B, auth: "x" },
    ]);

    expect(result.omitted).toEqual(
      expect.arrayContaining([
        { id: A, reason: "unchanged" },
        { id: B, reason: "notPasted" },
      ]),
    );
    expect(result.items).toEqual([]);
  });

  it("hands back a link whose revision the caller does not already have", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());

    const result = await store.fetchLinks([{ id: A, auth: "x" }]);

    expect(result.items).toEqual([
      { id: A, rev: R(1), blob: bytes(2, R(1) % 250) },
    ]);
    expect(result.omitted).toEqual([]);
  });
});

describe("LocalStore.deleteInbox and durable()", () => {
  it("deleteInbox removes a held row and is a silent no-op on one it never held", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([], [A]), new Set());
    expect((await store.listInbox("ws", "x")).map((row) => row.id)).toEqual([
      A,
    ]);

    await store.deleteInbox("ws", "x", A);
    await expect(store.deleteInbox("ws", "x", A)).resolves.toBeUndefined();

    expect(await store.listInbox("ws", "x")).toEqual([]);
  });

  it("durable() reflects the persistence port it was opened with", async () => {
    const store = await LocalStore.open(memoryPersistence());
    expect(store.durable()).toBe(false);
  });
});
