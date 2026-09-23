// src/link/transport-setting.test.ts
// Local by default; a device that already holds a link key keeps the relay.
import { describe, expect, it } from "vitest";
import {
  loadTransport,
  saveTransport,
  TRANSPORT_STORAGE_KEY,
} from "./transport-setting";
import { WORKSPACE_STORAGE_KEY, type KeyStore } from "./workspace";

function store(seed: Record<string, string> = {}): KeyStore {
  const map = new Map(Object.entries(seed));
  return {
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => void map.set(key, value),
    remove: async (key) => void map.delete(key),
  };
}

describe("loadTransport", () => {
  it("is local on a new install", async () => {
    expect(await loadTransport(store())).toBe("local");
  });
  it("keeps the relay where a link key is already stored", async () => {
    expect(await loadTransport(store({ [WORKSPACE_STORAGE_KEY]: "k" }))).toBe(
      "relay",
    );
  });
  it("follows what was saved, whatever the key says", async () => {
    const keys = store({ [WORKSPACE_STORAGE_KEY]: "k" });
    await saveTransport(keys, "local");
    expect(await keys.get(TRANSPORT_STORAGE_KEY)).toBe("local");
    expect(await loadTransport(keys)).toBe("local");
  });
  it("treats an unreadable store as a new install", async () => {
    const broken: KeyStore = {
      get: async () => {
        throw new Error("denied");
      },
      set: async () => undefined,
      remove: async () => undefined,
    };
    expect(await loadTransport(broken)).toBe("local");
  });
});
