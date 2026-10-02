// Pass-2, the v2.9.11 follow-up: a FRESH session (trustedThisSession
// empty - an upgrade, a reopened unsaved workbook, a Save As copy) can
// reissue a local rev the deck's LocalStore already holds, dropping fresh
// content silently. Fix: an untrusted floor is raised to the clock once.

import fc from "fast-check";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { isLocalRev, LOCAL_REV_BASE, localFloor } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import { LocalStore } from "../../src/link/local-store";
import { memoryPersistence } from "../../src/link/local-persist";
import { REGISTRY_SETTING } from "../../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

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

let links: typeof LinksModule;
let helpers: FakeHelpers;
let ws: Workspace;
let clockMs: number;

// Every step here - an upgrade, a close and reopen, a switch to editing a
// different file - is a real, human-paced action: seconds at least, never
// the same millisecond as the step before it.
function tick(ms = 5000): void {
  clockMs += ms;
  vi.setSystemTime(clockMs);
}

// A fresh pane session against the SAME workbook file: the fake host (the
// file's own state - cells, settings) is untouched, but every module,
// link-record.ts's trustedThisSession included, is reloaded cold - exactly
// what an add-in reload after an upgrade, a reopened workbook's pane, or a
// Save As copy's own separate pane all amount to.
async function freshSession(): Promise<typeof LinksModule> {
  vi.resetModules();
  return import("../../src/excel/links");
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  vi.useFakeTimers();
  clockMs = new Date("2026-01-01T00:00:00.000Z").getTime();
  vi.setSystemTime(clockMs);
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
  helpers.seed("Model!B4", [[0]]);
  helpers.select("Model!B4");
});
afterEach(() => {
  uninstallFakeHost();
  vi.useRealTimers();
});

// One local push of the current selection through the CURRENT session's
// links module, returning the bundle link row a real Copy press would have
// produced.
async function pushLocal(
  activeLinks: typeof LinksModule,
  value: number,
): Promise<{ id: string; rev: number; blob: Uint8Array }> {
  helpers.seed("Model!B4", [[value]]);
  const collector = new LocalCollector();
  const summary = await activeLinks.pushLinks("all", collector);
  expect(summary.failed).toBe(0);
  const link = collector.bundle().links[0]!;
  return { id: link.id, rev: link.rev, blob: link.blob };
}

function registryText(): string {
  const text = helpers.setting(REGISTRY_SETTING);
  if (text === null) throw new Error("no registry setting yet");
  return text;
}

async function freshStoreWith(
  ...pushes: { id: string; rev: number; blob: Uint8Array }[]
): Promise<LocalStore> {
  const store = await LocalStore.open(memoryPersistence());
  for (const push of pushes) {
    await store.ingest(
      {
        links: [
          { id: push.id, rev: push.rev, sentAt: push.rev, blob: push.blob },
        ],
        inbox: [],
      },
      new Set(),
    );
  }
  return store;
}

async function heldBlob(store: LocalStore, id: string): Promise<Uint8Array> {
  const held = await store.getLink(id, "auth");
  expect(held).not.toBe("unchanged");
  return (held as { blob: Uint8Array }).blob;
}

describe("legacy registry: a relay push left no localRev field at all", () => {
  it("swallows the first local push after the upgrade without the fix, but not with it", async () => {
    const { id } = await links.exportSelectionAsText(ws, new LocalCollector());
    tick();
    const first = await pushLocal(links, 1); // rev BASE+2 (export was BASE+1)
    const store = await freshStoreWith(first);
    tick();

    // A relay push reset `rev` to a small number, and this JSON predates
    // entry.localRev, so it was never written.
    const registry: { links: Record<string, unknown>[] } =
      JSON.parse(registryText());
    delete registry.links[0]!.localRev;
    registry.links[0]!.rev = 1;
    helpers.setSetting(REGISTRY_SETTING, JSON.stringify(registry));
    tick();

    // The upgrade: the add-in reloads, trustedThisSession starts empty.
    const afterUpgrade = await freshSession();
    const second = await pushLocal(afterUpgrade, 2);
    await store.ingest(
      {
        links: [
          {
            id: second.id,
            rev: second.rev,
            sentAt: second.rev,
            blob: second.blob,
          },
        ],
        inbox: [],
      },
      new Set(),
    );

    expect(await heldBlob(store, id)).toEqual(second.blob);
  });
});

describe("an unsaved workbook reopened: settings revert to what was last saved", () => {
  it("swallows the next local push without the fix, but not with it", async () => {
    const { id } = await links.exportSelectionAsText(ws, new LocalCollector());
    const savedSnapshot = registryText(); // "saved" at BASE+1
    tick();
    const unsaved = await pushLocal(links, 1); // BASE+2, never saved
    const store = await freshStoreWith(unsaved);
    tick();

    // Close without saving, reopen: settings revert to the last SAVE (the
    // unsaved push above is lost from the file, though the deck already has
    // it), and the reopened pane is a fresh session.
    helpers.setSetting(REGISTRY_SETTING, savedSnapshot);
    tick();
    const reopened = await freshSession();

    const afterReopen = await pushLocal(reopened, 2);
    await store.ingest(
      {
        links: [
          {
            id: afterReopen.id,
            rev: afterReopen.rev,
            sentAt: afterReopen.rev,
            blob: afterReopen.blob,
          },
        ],
        inbox: [],
      },
      new Set(),
    );

    expect(await heldBlob(store, id)).toEqual(afterReopen.blob);
  });
});

describe("a Save As copy: frozen at the moment of duplication", () => {
  it("swallows a push from the copy's own pane once the original has gone further, but not with the fix", async () => {
    const { id } = await links.exportSelectionAsText(ws, new LocalCollector());
    const forkPoint = registryText(); // the Save As copy's own registry, frozen here
    tick();
    const originalNext = await pushLocal(links, 1); // the original keeps going: BASE+2
    tick();
    const originalFurther = await pushLocal(links, 2); // BASE+3
    const store = await freshStoreWith(originalNext, originalFurther);
    tick();

    // Switch to the Save As copy: its registry is still the fork point, and
    // its pane is a separate session that never touched this link before.
    helpers.setSetting(REGISTRY_SETTING, forkPoint);
    tick();
    const copyPane = await freshSession();

    const fromCopy = await pushLocal(copyPane, 9);
    await store.ingest(
      {
        links: [
          {
            id: fromCopy.id,
            rev: fromCopy.rev,
            sentAt: fromCopy.rev,
            blob: fromCopy.blob,
          },
        ],
        inbox: [],
      },
      new Set(),
    );

    expect(await heldBlob(store, id)).toEqual(fromCopy.blob);
  });
});

describe("within one session, a healthy chain still climbs by exactly one push at a time", () => {
  it("never engages the clock once a link's floor has been trusted this session", async () => {
    await links.exportSelectionAsText(ws, new LocalCollector());
    tick();
    const a = await pushLocal(links, 1);
    tick();
    const b = await pushLocal(links, 2);
    tick();
    const c = await pushLocal(links, 3);

    expect(a.rev).toBeLessThan(b.rev);
    expect(b.rev).toBeLessThan(c.rev);
    expect(c.rev - b.rev).toBe(1);
    expect(b.rev - a.rev).toBe(1);
  });
});

describe("localFloor: the pure contract in isolation", () => {
  const SEED = 20260927;
  const RUNS = 300;

  it("a fresh floor (0) always passes through, trusted or not", () => {
    expect(localFloor(0, false)).toBe(0);
    expect(localFloor(0, true)).toBe(0);
  });

  it("a trusted floor always passes through unchanged, for any positive value", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: Number.MAX_SAFE_INTEGER }),
        (floor) => {
          expect(localFloor(floor, true)).toBe(floor);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("an untrusted, non-zero floor is always raised to a local rev above it", () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: LOCAL_REV_BASE }), (floor) => {
        const raised = localFloor(floor, false);
        expect(raised).toBeGreaterThan(floor);
        expect(isLocalRev(raised)).toBe(true);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("an untrusted floor already above LOCAL_REV_BASE still only ever climbs, never drops", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: LOCAL_REV_BASE + 1, max: Number.MAX_SAFE_INTEGER }),
        (floor) => {
          expect(localFloor(floor, false)).toBeGreaterThanOrEqual(floor);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
