import { describe, expect, it } from "vitest";
import {
  createWorkspace,
  forgetWorkspace,
  importWorkspace,
  loadWorkspace,
  officeKeyStore,
  type KeyStore,
} from "./workspace";

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
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

describe("workspace", () => {
  it("is absent until created, then persists", async () => {
    const store = memoryStore();
    expect(await loadWorkspace(store)).toBeNull();
    const created = await createWorkspace(store);
    const loaded = await loadWorkspace(store);
    expect(loaded?.id).toBe(created.id);
    expect(loaded?.auth).toBe(created.auth);
    expect(created.exportKey).toMatch(/^[A-Za-z0-9_-]{43}$/);
  });
  it("imports the same key on another host and derives the same id", async () => {
    const excel = memoryStore();
    const ppt = memoryStore();
    const created = await createWorkspace(excel);
    const imported = await importWorkspace(ppt, created.exportKey);
    expect(imported.id).toBe(created.id);
    expect(imported.enc).toEqual(created.enc);
    await expect(importWorkspace(ppt, "short")).rejects.toThrow(/link key/);
  });
  it("forgets", async () => {
    const store = memoryStore();
    await createWorkspace(store);
    await forgetWorkspace(store);
    expect(await loadWorkspace(store)).toBeNull();
  });
});

describe("officeKeyStore", () => {
  it("falls back to a working store when no host storage exists", async () => {
    const store = officeKeyStore();
    const created = await createWorkspace(store);
    expect((await loadWorkspace(store))?.id).toBe(created.id);
    await forgetWorkspace(store);
    expect(await loadWorkspace(store)).toBeNull();
  });
});
