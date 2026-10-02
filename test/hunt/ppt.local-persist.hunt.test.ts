// Hunt pass 1: src/link/local-persist.ts's IndexedDB failure paths - open
// erroring or blocked, a read or write failing, and a synchronous throw from
// db.transaction() itself. Every one must degrade `durable` to false and
// resolve, never throw into a paste (fake-indexeddb has no dial for this).

import { IDBFactory } from "fake-indexeddb";
import { describe, expect, it } from "vitest";
import { openLocalPersistence } from "../../src/link/local-persist";

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

// A settable open() request: resolve(db) fires onsuccess, reject() fires
// onerror, block() fires onblocked - whichever local-persist.ts's openDb
// wires up, next microtask, the way a real async IndexedDB request would.
function openRequest(): {
  request: IDBOpenDBRequest;
  resolve: (db: IDBDatabase) => void;
  reject: (error?: Error) => void;
  block: () => void;
} {
  const request = {} as IDBOpenDBRequest;
  return {
    request,
    resolve: (db) => {
      (request as { result?: IDBDatabase }).result = db;
      queueMicrotask(() => request.onsuccess?.(new Event("success")));
    },
    reject: (error) => {
      (request as { error?: Error | null }).error = error ?? null;
      queueMicrotask(() => request.onerror?.(new Event("error")));
    },
    block: () => {
      // openDb's own onblocked handler reads no event field, so a plain
      // Event stands in for the IDBVersionChangeEvent the real API fires.
      queueMicrotask(() =>
        request.onblocked?.(new Event("blocked") as IDBVersionChangeEvent),
      );
    },
  };
}

// A fake IDBDatabase whose transaction() either throws synchronously (a
// closing connection) or returns a transaction this test drives by hand:
// objectStore().getAll()/put()/delete()/clear() succeed at once, and the
// caller decides whether the transaction itself completes or aborts.
function fakeDb(options: {
  transactionThrows?: boolean;
  getAllFails?: boolean;
  transactionOutcome?: "complete" | "abort" | "error";
}): IDBDatabase {
  return {
    transaction() {
      if (options.transactionThrows) throw new Error("connection is closing");
      // One object, mutated in place: write() assigns oncomplete/onerror/
      // onabort onto exactly the reference this factory hands back, and the
      // queued microtask below must fire those same three properties, not a
      // copy taken before write() ever sets them.
      const tx = {
        objectStore: () => ({
          put: () => undefined,
          delete: () => undefined,
          clear: () => undefined,
          getAll: () => {
            const request = {} as IDBRequest;
            queueMicrotask(() => {
              if (options.getAllFails) {
                (request as { error?: Error | null }).error = new Error(
                  "read failed",
                );
                request.onerror?.(new Event("error"));
              } else {
                (request as { result?: unknown }).result = [];
                request.onsuccess?.(new Event("success"));
              }
            });
            return request;
          },
        }),
      } as unknown as IDBTransaction;
      queueMicrotask(() => {
        const outcome = options.transactionOutcome ?? "complete";
        if (outcome === "complete") tx.oncomplete?.(new Event("complete"));
        else if (outcome === "abort") tx.onabort?.(new Event("abort"));
        else tx.onerror?.(new Event("error"));
      });
      return tx;
    },
  } as unknown as IDBDatabase;
}

function factoryThatOpens(db: IDBDatabase): IDBFactory {
  return {
    open: () => {
      const { request, resolve } = openRequest();
      resolve(db);
      return request;
    },
  } as unknown as IDBFactory;
}

describe("openLocalPersistence: the open call itself fails", () => {
  it("degrades to memory when open() answers onerror", async () => {
    const factory = {
      open: () => {
        const { request, reject } = openRequest();
        reject(new Error("SecurityError"));
        return request;
      },
    } as unknown as IDBFactory;

    const persist = await openLocalPersistence(factory);

    expect(persist.durable).toBe(false);
    await expect(persist.load()).resolves.toEqual({ links: [], inbox: [] });
  });

  it("degrades to memory when open() answers onblocked", async () => {
    const factory = {
      open: () => {
        const { request, block } = openRequest();
        block();
        return request;
      },
    } as unknown as IDBFactory;

    const persist = await openLocalPersistence(factory);

    expect(persist.durable).toBe(false);
  });
});

describe("openLocalPersistence: a read fails after a good open", () => {
  it("load() degrades durable to false and answers empty, never throwing", async () => {
    const factory = factoryThatOpens(fakeDb({ getAllFails: true }));
    const persist = await openLocalPersistence(factory);
    expect(persist.durable).toBe(true);

    await expect(persist.load()).resolves.toEqual({ links: [], inbox: [] });

    expect(persist.durable).toBe(false);
  });
});

describe("openLocalPersistence: a write fails after a good open", () => {
  it("a transaction that aborts turns durable false and still resolves", async () => {
    const factory = factoryThatOpens(fakeDb({ transactionOutcome: "abort" }));
    const persist = await openLocalPersistence(factory);

    await expect(persist.putLinks([link])).resolves.toBeUndefined();

    expect(persist.durable).toBe(false);
  });

  it("a transaction that errors turns durable false and still resolves", async () => {
    const factory = factoryThatOpens(fakeDb({ transactionOutcome: "error" }));
    const persist = await openLocalPersistence(factory);

    await expect(persist.putInbox([row])).resolves.toBeUndefined();

    expect(persist.durable).toBe(false);
  });

  it("a closing connection throws synchronously from transaction() and still resolves", async () => {
    const factory = factoryThatOpens(fakeDb({ transactionThrows: true }));
    const persist = await openLocalPersistence(factory);

    await expect(persist.deleteLinks([link.id])).resolves.toBeUndefined();

    expect(persist.durable).toBe(false);
  });
});

describe("openLocalPersistence: deleteInbox on the real (fake-indexeddb) path", () => {
  it("removes an inbox row and leaves links untouched", async () => {
    const factory = new IDBFactory();
    const persist = await openLocalPersistence(factory);
    await persist.putLinks([link]);
    await persist.putInbox([row]);

    await persist.deleteInbox([row.id]);

    const loaded = await persist.load();
    expect(loaded.inbox).toEqual([]);
    expect(loaded.links).toEqual([link]);
  });
});
