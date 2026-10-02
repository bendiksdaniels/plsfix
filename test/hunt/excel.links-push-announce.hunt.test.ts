// Attacks pushRegistry/pushOne (src/excel/links.ts): whether a push that is
// reported as "failed" ever still advances the registry's revision or its
// local floor. Local mode's Copy selected/Copy all is the one caller that
// sets `options.announce`, and always pairs it with a LocalCollector.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { LOCAL_REV_BASE } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
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
  type FakeWorkbook,
} from "../fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let ws: Workspace;

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

function registryEntry(id: string): {
  rev: number;
  localRev?: number;
  lastPushedAt: string | null;
} {
  const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING))) as {
    links: {
      id: string;
      rev: number;
      localRev?: number;
      lastPushedAt: string | null;
    }[];
  };
  return registry.links.find((link) => link.id === id)!;
}

function registryRev(id: string): number {
  return registryEntry(id).rev;
}

function registryLocalRev(id: string): number | undefined {
  return registryEntry(id).localRev;
}

function registryLastPushedAt(id: string): string | null {
  return registryEntry(id).lastPushedAt;
}

// The shape production actually builds (src/pane/links-transport.ts's
// copy() makes one fresh LocalCollector per press and hands it to pushLinks
// as BOTH the relay and options.announce's partner - never a real relay: a
// relay-mode push never sets `announce` at all, src/pane/links-tab.ts:408).
// A fresh collector every time, because that is what a fresh press builds.
function failingAnnounce(): LocalCollector {
  const collector = new LocalCollector();
  collector.postInbox = () => Promise.reject(new Error("inbox is full"));
  return collector;
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.select("Model!B4:F5");
});
afterEach(() => uninstallFakeHost());

describe("pushRegistry with an announce that fails after the payload already landed", () => {
  it("does not advance the registry's revision or lastPushedAt for a push it reports as failed, though the local floor moves underneath it", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    const before = registryLastPushedAt(id);
    const R = registryRev(id);
    expect(R).toBe(LOCAL_REV_BASE + 1);
    expect(registryLocalRev(id)).toBe(R);

    // Change the cell so a real push has something new to send, then make
    // the announce step (not the payload push) refuse on a fresh collector
    // - exactly what one press of Copy selected builds.
    helpers.seed("Model!B4", [[99]]);
    const summary = await links.pushLinks("all", failingAnnounce(), {
      announce: ws,
    });

    expect(summary).toEqual({
      pushed: 0,
      missing: 0,
      failed: 1,
      failures: [`Model!B4:F5: inbox is full`],
    });
    // The user was told this push failed: the registry's own rev must
    // agree, not quietly show a newer revision than what was reported - but
    // the payload push (before the announce) already landed, so the local
    // floor pushPayload counts on from next time is left one ahead of it.
    expect(registryRev(id)).toBe(R);
    expect(registryLastPushedAt(id)).toBe(before);
    expect(registryLocalRev(id)).toBe(R + 1);
  });

  it("still advances the registry once the announce succeeds on a later retry, past the floor the failed attempt already moved", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    const R = registryRev(id);
    helpers.seed("Model!B4", [[99]]);
    await links.pushLinks("all", failingAnnounce(), { announce: ws });
    expect(registryRev(id)).toBe(R);
    expect(registryLocalRev(id)).toBe(R + 1);

    const retry = new LocalCollector();
    const summary = await links.pushLinks("all", retry, { announce: ws });

    expect(summary).toEqual({ pushed: 1, missing: 0, failed: 0, failures: [] });
    // The failed first attempt's payload already moved the local floor to
    // R + 1 (never rolled back), so this fully-successful retry - a fresh
    // press, a fresh collector - counts on from there and lands one past
    // it, never colliding with the stale attempt.
    expect(registryRev(id)).toBeGreaterThan(R + 1);
    expect(registryRev(id)).toBe(R + 2);
    expect(retry.bundle().inbox).toHaveLength(1);
  });

  it("never announces a link whose payload push itself failed", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    const R = registryRev(id);
    const collector = new LocalCollector();
    const announce = vi.fn(collector.postInbox.bind(collector));
    collector.postInbox = announce;
    collector.putLink = () => Promise.reject(new Error("boom"));

    const summary = await links.pushLinks("all", collector, { announce: ws });

    expect(summary.failed).toBe(1);
    expect(registryRev(id)).toBe(R);
    expect(announce).not.toHaveBeenCalled();
  });

  it("keeps failing the same way on three repeated attempts, never leaking a bumped registry revision, though the local floor climbs underneath it", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    const R = registryRev(id);

    for (let attempt = 0; attempt < 3; attempt += 1) {
      helpers.seed("Model!B4", [[100 + attempt]]);
      const summary = await links.pushLinks("all", failingAnnounce(), {
        announce: ws,
      });
      expect(summary.failed).toBe(1);
      // The registry - what the modeller can inspect - must keep agreeing
      // with "failed" on every repeat, even though the local floor
      // pushPayload counts on from (never rolled back once it lands, the
      // same trade-off publish()'s own rollback accepts) climbs by one
      // every attempt underneath it.
      expect(registryRev(id)).toBe(R);
      expect(registryLocalRev(id)).toBe(R + attempt + 1);
    }
  });
});
