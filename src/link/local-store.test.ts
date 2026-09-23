// src/link/local-store.test.ts
// The store's own rules: no downgrade, the Inbox skips what the deck holds,
// the budget evicts the oldest paste, a reopen brings everything back.
import { describe, expect, it } from "vitest";
import type { Bundle } from "./bundle";
import { LOCAL_REV_BASE } from "./local";
import {
  memoryPersistence,
  type LocalPersistence,
  type StoredLocalInbox,
  type StoredLocalLink,
} from "./local-persist";
import { LocalStore, NOT_PASTED_HERE } from "./local-store";

const A = "a".repeat(32);
const B = "b".repeat(32);
const R = (n: number): number => LOCAL_REV_BASE + n;
const bytes = (n: number, fill = 1): Uint8Array => new Uint8Array(n).fill(fill);

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

// Keeps what it is given, so a second LocalStore.open reads it back.
function mapPersistence(): LocalPersistence {
  const links = new Map<string, StoredLocalLink>();
  const inbox = new Map<string, StoredLocalInbox>();
  return {
    durable: true,
    load: async () => ({
      links: [...links.values()],
      inbox: [...inbox.values()],
    }),
    putLinks: async (list) => list.forEach((l) => links.set(l.id, l)),
    putInbox: async (list) => list.forEach((r) => inbox.set(r.id, r)),
    deleteLinks: async (ids) => ids.forEach((id) => links.delete(id)),
    deleteInbox: async (ids) => ids.forEach((id) => inbox.delete(id)),
    clear: async () => {
      links.clear();
      inbox.clear();
    },
  };
}

describe("LocalStore.ingest", () => {
  it("never downgrades: an older bundle leaves the newest in place", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(3), 2]]), new Set());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());
    const [status] = await store.status([{ id: A, auth: "x" }]);
    expect(status?.rev).toBe(R(3));
    await expect(store.getLinkRev(A, "x", R(1))).resolves.toEqual({
      rev: R(1),
      blob: bytes(2, R(1) % 250),
    });
  });

  it("keeps inbox rows only for links the deck does not hold", async () => {
    const store = await LocalStore.open(memoryPersistence());
    const result = await store.ingest(
      bundle(
        [
          [A, R(1), 2],
          [B, R(1), 2],
        ],
        [A, B],
      ),
      new Set([A]),
    );
    expect(result).toEqual({ linkIds: [A, B], waiting: 1 });
    expect((await store.listInbox("ws", "x")).map((row) => row.id)).toEqual([
      B,
    ]);
  });

  it("evicts the oldest paste first, never the one in progress", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), {
      now: () => (now += 1),
      budget: 10,
    });
    await store.ingest(bundle([[A, R(1), 6]]), new Set());
    await store.ingest(bundle([[B, R(1), 6]]), new Set());
    const [a, b] = await store.status([
      { id: A, auth: "x" },
      { id: B, auth: "x" },
    ]);
    expect(a?.rev).toBeNull();
    expect(b?.rev).toBe(R(1));
  });

  it("comes back from its persistence after a reopen, and clear empties both", async () => {
    const persist = mapPersistence();
    const first = await LocalStore.open(persist);
    await first.ingest(bundle([[A, R(2), 2]], [A]), new Set());
    const second = await LocalStore.open(persist);
    expect((await second.status([{ id: A, auth: "x" }]))[0]?.rev).toBe(R(2));
    expect(await second.listInbox("ws", "x")).toHaveLength(1);
    await second.clear();
    expect(
      (await (await LocalStore.open(persist)).status([{ id: A, auth: "x" }]))[0]
        ?.rev,
    ).toBeNull();
  });

  it("says in a sentence that a link was never pasted here", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await expect(store.getLink(A, "x")).rejects.toThrow(NOT_PASTED_HERE);
  });
});
