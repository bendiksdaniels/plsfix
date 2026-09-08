// The relay client's edges, over a stubbed fetch: the two deletes (never
// driven by any other suite, because the integration tests all run against the
// fake relay), the status codes the routes can answer with, and the host key
// store the pairing is remembered in.
// Invariant: every call ends as a value or a typed RelayError, never a raw
// throw and never a refusal a caller could not act on.
import { describe, expect, it } from "vitest";
import { isRelayError, RelayError } from "./relay";
import { client, rejection } from "./relay.support";
import {
  officeKeyStore,
  deriveWorkspace,
  loadWorkspace,
  createWorkspace,
  importWorkspace,
  forgetWorkspace,
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
} from "./workspace";

const ID = "a".repeat(32);
const WS = "W".repeat(43);

describe("RelayClient.deleteLink", () => {
  it("sends the bearer and takes 200 or 204 for done", async () => {
    for (const status of [200, 204]) {
      const { relay, calls } = client(() => new Response(null, { status }));
      await expect(relay.deleteLink(ID, "AUTH")).resolves.toBeUndefined();
      expect(calls[0]!.url).toBe(`https://x.test/modelis/api/links/${ID}`);
      expect(calls[0]!.init.method).toBe("DELETE");
      expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(
        "Bearer AUTH",
      );
    }
  });

  // The Excel side branches on this one: a link the relay never had, or one a
  // sweep took, is "already removed", not a failed removal.
  it("still reports a link the relay does not hold as missing", async () => {
    const { relay } = client(
      () =>
        new Response(JSON.stringify({ error: "not found" }), { status: 404 }),
    );
    const error = await rejection(relay.deleteLink(ID, "AUTH"));
    expect(isRelayError(error) && error.kind).toBe("missing");
    expect(String(error)).toBe(
      `RelayError: relay DELETE /modelis/api/links/${ID}: 404 not found`,
    );
  });
});

describe("RelayClient.deleteInbox", () => {
  it("sends the bearer on the workspace path", async () => {
    const { relay, calls } = client(() => new Response(null, { status: 200 }));
    await relay.deleteInbox(WS, "AUTH", ID);
    expect(calls[0]!.url).toBe(`https://x.test/modelis/api/inbox/${WS}/${ID}`);
    expect(calls[0]!.init.method).toBe("DELETE");
    expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(
      "Bearer AUTH",
    );
  });

  // The relay answers 404 both for a row that is already gone (the item was
  // taken, or its seven-day TTL ran out under an open deck) and for one this
  // key never wrote - and it cannot tell the two apart on purpose. Neither is
  // something the deck can act on, and the delete runs after the shape is
  // already on the slide, so a rejection here would report a failure for an
  // insert that worked and invite a second one.
  it("takes 404 for done, because the row is gone either way", async () => {
    const { relay, calls } = client(
      () =>
        new Response(JSON.stringify({ error: "not found" }), { status: 404 }),
    );
    await expect(relay.deleteInbox(WS, "AUTH", ID)).resolves.toBeUndefined();
    expect(calls).toHaveLength(1);
  });

  it("still reports a refusal that is not the row being gone", async () => {
    for (const [status, kind] of [
      [401, "auth"],
      [403, "auth"],
      [500, "server"],
    ] as const) {
      const { relay } = client(() => new Response("", { status }));
      const error = await rejection(relay.deleteInbox(WS, "AUTH", ID));
      expect(isRelayError(error) && error.kind, String(status)).toBe(kind);
    }
  });
});

describe("RelayClient status mapping", () => {
  // 401 is the relay's answer to a request with no bearer at all, 409 and 507
  // are the two an intermediary or the store can add: none of them may reach a
  // caller as anything but a kind it can branch on.
  it("maps every status the relay and its gateway can answer with", async () => {
    for (const [status, kind] of [
      [401, "auth"],
      [403, "auth"],
      [404, "missing"],
      [409, "server"],
      [413, "tooLarge"],
      [429, "server"],
      [507, "server"],
      [502, "server"],
    ] as const) {
      const { relay } = client(() => new Response("", { status }));
      const error = await rejection(
        relay.putLink(ID, "A", new Uint8Array([1])),
      );
      expect(isRelayError(error) && error.kind, String(status)).toBe(kind);
      expect(isRelayError(error) && error.status).toBe(status);
    }
  });

  // The relay names the wait in a Retry-After header. The client does not act
  // on it - nothing here retries by itself - but the status travels on the
  // error, so a caller that wants to can tell 429 from any other server fault.
  it("carries the status of a rate refusal beside its reason", async () => {
    const { relay } = client(
      () =>
        new Response(JSON.stringify({ error: "too many requests" }), {
          status: 429,
          headers: { "Retry-After": "3" },
        }),
    );
    const error = await rejection(relay.putLink(ID, "A", new Uint8Array([1])));
    expect(isRelayError(error) && error.status).toBe(429);
    expect(String(error)).toContain("429 too many requests");
  });

  it("is a network error, not a throw, when the relay cannot be reached", async () => {
    const { relay } = client(() => {
      throw new TypeError("Failed to fetch");
    });
    const error = await rejection(relay.listInbox(WS, "AUTH"));
    expect(isRelayError(error) && error.kind).toBe("network");
    expect(error).toBeInstanceOf(RelayError);
  });
});

// The pairing key's store is the one piece of this folder that touches a host
// API. Three hosts answer it three ways, and only the in-memory stand-in has
// ever run in a test.
describe("officeKeyStore", () => {
  interface Hosted {
    OfficeRuntime?: unknown;
    localStorage?: unknown;
  }
  const host = globalThis as unknown as Hosted;

  async function roundTrip(store: KeyStore): Promise<(string | null)[]> {
    const seen: (string | null)[] = [await store.get("k")];
    await store.set("k", "V");
    seen.push(await store.get("k"));
    await store.remove("k");
    seen.push(await store.get("k"));
    return seen;
  }

  it("uses OfficeRuntime.storage where the host has one", async () => {
    const map = new Map<string, string>();
    host.OfficeRuntime = {
      storage: {
        getItem: async (key: string) => map.get(key) ?? null,
        setItem: async (key: string, value: string) => {
          map.set(key, value);
        },
        removeItem: async (key: string) => {
          map.delete(key);
        },
      },
    };
    try {
      expect(await roundTrip(officeKeyStore())).toEqual([null, "V", null]);
    } finally {
      delete host.OfficeRuntime;
    }
  });

  // A read that fails is not "this machine has no key": the Excel tab shows the
  // reason and keeps Generate off, because generating a second key unpairs
  // every deck holding the first.
  it("lets a storage failure through rather than reading as unpaired", async () => {
    host.OfficeRuntime = {
      storage: {
        getItem: async () => {
          throw new Error("storage refused");
        },
        setItem: async () => undefined,
        removeItem: async () => undefined,
      },
    };
    try {
      await expect(loadWorkspace(officeKeyStore())).rejects.toThrow(
        "storage refused",
      );
    } finally {
      delete host.OfficeRuntime;
    }
  });

  it("falls back to localStorage, and to memory when that is refused", async () => {
    expect(await roundTrip(officeKeyStore())).toEqual([null, "V", null]);
    const real = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      get() {
        throw new Error("site data blocked");
      },
    });
    try {
      // The process-wide map, so both calls of one pane see one pairing.
      const store = officeKeyStore();
      await store.set(WORKSPACE_STORAGE_KEY, "kept");
      expect(await officeKeyStore().get(WORKSPACE_STORAGE_KEY)).toBe("kept");
      await store.remove(WORKSPACE_STORAGE_KEY);
    } finally {
      if (real) Object.defineProperty(globalThis, "localStorage", real);
      else delete host.localStorage;
    }
  });
});

describe("workspace lifecycle", () => {
  function memory(): KeyStore {
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

  it("remembers a generated key and forgets it again", async () => {
    const store = memory();
    expect(await loadWorkspace(store)).toBeNull();
    const made = await createWorkspace(store);
    expect((await loadWorkspace(store))?.exportKey).toBe(made.exportKey);
    await forgetWorkspace(store);
    expect(await loadWorkspace(store)).toBeNull();
  });

  // A stored value that is not 32 bytes is a key from somewhere else, not a
  // pairing: it reads as unpaired instead of deriving a workspace nobody holds.
  it("reads a truncated or foreign stored key as unpaired", async () => {
    for (const stored of ["", "abc", "not base64url!", "A".repeat(42)]) {
      const store = memory();
      await store.set(WORKSPACE_STORAGE_KEY, stored);
      expect(await loadWorkspace(store), stored).toBeNull();
    }
  });

  it("imports a key with surrounding whitespace and refuses a mistyped one", async () => {
    const store = memory();
    const made = await deriveWorkspace(new Uint8Array(32).fill(7));
    const back = await importWorkspace(store, ` ${made.exportKey}\n`);
    expect(back.id).toBe(made.id);
    await expect(importWorkspace(store, "nope")).rejects.toThrow(
      "invalid link key",
    );
    // The refused import left the pairing that was already there.
    expect((await loadWorkspace(store))?.id).toBe(made.id);
  });
});
