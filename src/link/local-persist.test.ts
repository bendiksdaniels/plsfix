// src/link/local-persist.test.ts
// IndexedDB behind the port: a round trip, deletes, clear, and memory when
// the webview has no IndexedDB or refuses it.
import { IDBFactory } from "fake-indexeddb";
import { describe, expect, it } from "vitest";
import { openLocalPersistence } from "./local-persist";

const link = {
  id: "a".repeat(32),
  rev: 5,
  sentAt: 1,
  blob: new Uint8Array([1, 2]),
  pastedAt: 2,
};
const row = {
  id: "b".repeat(32),
  createdAt: 3,
  blob: new Uint8Array([3]),
  pastedAt: 4,
};

describe("openLocalPersistence", () => {
  it("keeps links and inbox rows across a reopen", async () => {
    const factory = new IDBFactory();
    const first = await openLocalPersistence(factory);
    expect(first.durable).toBe(true);
    await first.putLinks([link]);
    await first.putInbox([row]);
    const loaded = await (await openLocalPersistence(factory)).load();
    expect(loaded.links).toEqual([link]);
    expect(loaded.inbox).toEqual([row]);
  });

  it("deletes and clears", async () => {
    const factory = new IDBFactory();
    const persist = await openLocalPersistence(factory);
    await persist.putLinks([link]);
    await persist.putInbox([row]);
    await persist.deleteLinks([link.id]);
    expect((await persist.load()).links).toEqual([]);
    await persist.clear();
    expect((await persist.load()).inbox).toEqual([]);
  });

  it("falls back to memory without IndexedDB or when opening fails", async () => {
    expect((await openLocalPersistence(undefined)).durable).toBe(false);
    const refusing = {
      open: () => {
        throw new Error("SecurityError");
      },
    } as unknown as IDBFactory;
    expect((await openLocalPersistence(refusing)).durable).toBe(false);
  });
});
