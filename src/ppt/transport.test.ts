// src/ppt/transport.test.ts
// PptTransport's own rules: which mode loadTransport hands back decides what
// relay()/workspace() answer, and the local store opens once - only once
// local mode is actually used - and is kept across later mode switches.

import { describe, expect, it } from "vitest";
import {
  memoryPersistence,
  type LocalPersistence,
} from "../link/local-persist";
import { LocalStore } from "../link/local-store";
import { TRANSPORT_STORAGE_KEY } from "../link/transport-setting";
import {
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
  type Workspace,
} from "../link/workspace";
import { FakeRelay } from "../../test/fakerelay";
import { openPptTransport } from "./transport";

function mapKeyStore(seed: Record<string, string> = {}): KeyStore {
  const map = new Map(Object.entries(seed));
  return {
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => {
      map.set(key, value);
    },
    remove: async (key) => {
      map.delete(key);
    },
  };
}

// Counts every open, the way the brief's "a persistence stub counts opens"
// asks for: LocalStore.open calls persist.load() exactly once per open.
function countingPersistence(): {
  persistence: () => Promise<LocalPersistence>;
  opens: () => number;
} {
  let opens = 0;
  return {
    persistence: async () => {
      const base = memoryPersistence();
      return {
        ...base,
        load: async () => {
          opens += 1;
          return base.load();
        },
      };
    },
    opens: () => opens,
  };
}

const PAIRED: Workspace = {
  id: "paired-id",
  enc: new Uint8Array(32).fill(1),
  auth: "paired-auth",
  exportKey: "paired-export-key",
};

describe("openPptTransport", () => {
  it("defaults local with no stored mode or key, and opens the store once", async () => {
    const remote = new FakeRelay();
    const counting = countingPersistence();
    const transport = await openPptTransport({
      keyStore: mapKeyStore(),
      remote,
      paired: () => null,
      persistence: counting.persistence,
    });
    expect(transport.mode()).toBe("local");
    expect(transport.relay()).toBeInstanceOf(LocalStore);
    expect(transport.store()).toBeInstanceOf(LocalStore);
    expect(counting.opens()).toBe(1);
  });

  it("defaults relay when a key is already stored, and opens no store", async () => {
    const remote = new FakeRelay();
    const counting = countingPersistence();
    const transport = await openPptTransport({
      keyStore: mapKeyStore({ [WORKSPACE_STORAGE_KEY]: "x" }),
      remote,
      paired: () => PAIRED,
      persistence: counting.persistence,
    });
    expect(transport.mode()).toBe("relay");
    expect(transport.relay()).toBe(remote);
    expect(transport.store()).toBeNull();
    expect(transport.workspace()).toBe(PAIRED);
    expect(counting.opens()).toBe(0);
  });

  it("honours an explicitly stored mode over the key-presence default", async () => {
    const remote = new FakeRelay();
    const transport = await openPptTransport({
      keyStore: mapKeyStore({
        [WORKSPACE_STORAGE_KEY]: "x",
        [TRANSPORT_STORAGE_KEY]: "local",
      }),
      remote,
      paired: () => PAIRED,
    });
    expect(transport.mode()).toBe("local");
  });

  it("answers the fixed local workspace in local mode, never the paired one", async () => {
    const remote = new FakeRelay();
    const transport = await openPptTransport({
      keyStore: mapKeyStore(),
      remote,
      paired: () => PAIRED,
    });
    const ws = transport.workspace();
    expect(ws).not.toBeNull();
    expect(ws?.id).not.toBe(PAIRED.id);
  });

  it("opens the store on first switch to local and keeps it across later switches", async () => {
    const remote = new FakeRelay();
    const counting = countingPersistence();
    const transport = await openPptTransport({
      keyStore: mapKeyStore({ [WORKSPACE_STORAGE_KEY]: "x" }),
      remote,
      paired: () => PAIRED,
      persistence: counting.persistence,
    });
    expect(transport.store()).toBeNull();

    await transport.setMode("local");
    expect(counting.opens()).toBe(1);
    const opened = transport.store();
    expect(opened).toBeInstanceOf(LocalStore);

    await transport.setMode("relay");
    expect(transport.relay()).toBe(remote);
    expect(transport.store()).toBe(opened);

    await transport.setMode("local");
    expect(counting.opens()).toBe(1);
    expect(transport.store()).toBe(opened);
  });

  it("persists the chosen mode through the key store", async () => {
    const remote = new FakeRelay();
    const keyStore = mapKeyStore({ [WORKSPACE_STORAGE_KEY]: "x" });
    const transport = await openPptTransport({
      keyStore,
      remote,
      paired: () => PAIRED,
    });
    expect(transport.mode()).toBe("relay");

    await transport.setMode("local");
    expect(await keyStore.get(TRANSPORT_STORAGE_KEY)).toBe("local");
  });
});
