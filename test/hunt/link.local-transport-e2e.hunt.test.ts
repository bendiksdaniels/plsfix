// End-to-end local transport lens over src/excel/links.ts + LocalCollector +
// the registry. Pins a real bug (first describe block): local -> relay ->
// local collided the second local push's revision with the first's, leaving
// a deck stale for as many copies as were made before the relay push.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { deriveLinkKeys, open } from "../../src/link/crypto";
import {
  anchorName,
  decodePayload,
  decodeRegistry,
  encodeRegistry,
  REGISTRY_SETTING,
} from "../../src/link/model";
import { LOCAL_REV_BASE } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let relay: FakeRelay;
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

function registry(): {
  links: { id: string; rev: number; localRev?: number }[];
} {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  host.workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
  helpers.seed("Model!B4", [[1]]);
  helpers.select("Model!B4");
});
afterEach(() => uninstallFakeHost());

describe("the transport setting switched local -> relay -> local", () => {
  it("never reproduces an earlier local revision, even though the registry's own rev was reset by the relay push in between", async () => {
    // t1: a fresh local export. As text, not the plain range export every
    // other test here uses: the fake host's getImage() is signature-only
    // (fakepng.ts), sized from the selection and blind to the cell's own
    // value, so a picture link could never prove t3's content differs from
    // t1's the way a rendered text link can.
    const first = new LocalCollector();
    const { id } = await links.exportSelectionAsText(ws, first);
    const firstBlob = first.bundle().links[0]!.blob;
    const token = decodeRegistry(helpers.setting(REGISTRY_SETTING)).links[0]!
      .token;
    expect(registry().links[0]!.rev).toBe(LOCAL_REV_BASE + 1);

    // t2: switch to relay, push the changed content. The relay has never
    // seen this id, so it starts its own count at 1 - small, and unrelated
    // to the local space.
    helpers.seed("Model!B4", [[2]]);
    await links.pushLinks("all", relay);
    expect(registry().links[0]!.rev).toBe(1);

    // t3: switch back to local, push the changed content again.
    helpers.seed("Model!B4", [[3]]);
    const second = new LocalCollector();
    await links.pushLinks("all", second);
    const secondLink = second.bundle().links[0]!;

    // The bug: without the fix, secondLink.rev === LOCAL_REV_BASE + 1 again
    // (nextLocalRev(1) === nextLocalRev(anything <= LOCAL_REV_BASE)), which
    // src/link/local-store.ts's mergeRevision (rev === held.rev) reads as
    // no-op and keeps t1's stale blob - for as many further copies as were
    // made before the relay push - so a deck that already pasted t1 would
    // not see t3's content.
    expect(secondLink.rev).toBeGreaterThan(LOCAL_REV_BASE + 1);
    // Not a blob comparison: seal()'s IV is random, so two pushes of the
    // SAME content would already produce different ciphertext. The hash is
    // content, never the envelope (payloadOf's own invariant in
    // src/excel/link-record.ts) - proving it changed proves t3's cell value
    // actually reached the bundle, not t1's.
    const keys = await deriveLinkKeys(token);
    const firstPayload = decodePayload(await open(keys.enc, id, firstBlob));
    const secondPayload = decodePayload(
      await open(keys.enc, id, secondLink.blob),
    );
    expect(secondPayload.hash).not.toBe(firstPayload.hash);
    expect(registry().links[0]).toMatchObject({ id, localRev: secondLink.rev });
  });

  it("keeps climbing across three round trips, not just one", async () => {
    await links.exportSelection(ws, relay); // establishes the link on the relay first
    const localRevs: number[] = [];

    for (let round = 0; round < 3; round += 1) {
      helpers.seed("Model!B4", [[100 + round]]);
      const local = new LocalCollector();
      await links.pushLinks("all", local);
      localRevs.push(local.bundle().links[0]!.rev);

      helpers.seed("Model!B4", [[200 + round]]);
      await links.pushLinks("all", relay); // resets the registry's rev to a small relay number again
    }

    expect(localRevs).toEqual([...localRevs].sort((a, b) => a - b));
    expect(new Set(localRevs).size).toBe(3); // three distinct revisions, never repeated
  });
});

describe("registry round-trip: localRev survives encode/decode", () => {
  it("keeps localRev across a write and a fresh read", async () => {
    const collector = new LocalCollector();
    await links.exportSelection(ws, collector);
    const before = registry().links[0]!;
    expect(before.localRev).toBe(LOCAL_REV_BASE + 1);

    const decoded = decodeRegistry(helpers.setting(REGISTRY_SETTING));
    expect(decoded.links[0]!.localRev).toBe(LOCAL_REV_BASE + 1);
    // encodeRegistry -> decodeRegistry must not drop it either.
    const roundTripped = decodeRegistry(encodeRegistry(decoded));
    expect(roundTripped.links[0]!.localRev).toBe(LOCAL_REV_BASE + 1);
  });

  it("never appears at all for a link that was only ever pushed through the relay", async () => {
    await links.exportSelection(ws, relay);
    expect(registry().links[0]!.localRev).toBeUndefined();
  });
});

describe("repeats: the same link copied three times in a row", () => {
  it("increases the revision each time and never duplicates a bundle row within one collector", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    const shared = new LocalCollector();

    for (let attempt = 0; attempt < 3; attempt += 1) {
      helpers.seed("Model!B4", [[100 + attempt]]);
      await links.pushLinks("all", shared, { announce: ws });
    }

    const bundle = shared.bundle();
    expect(bundle.links).toHaveLength(1); // never three stacked rows
    expect(bundle.links[0]!.rev).toBe(LOCAL_REV_BASE + 4); // +1 export, +3 pushes
    expect(bundle.inbox).toHaveLength(1);
    expect(bundle.inbox[0]!.id).toBe(id);
  });
});

describe("a moved source, pushed locally", () => {
  it("still resolves through the anchor and copies the content at the new address", async () => {
    const { id } = await links.exportSelection(ws, new LocalCollector());
    // Excel rewrites a hidden name's formula when rows move under it; the
    // fake models that directly, the same way
    // links.integration.test.ts's own "renders through the anchor name, not
    // the original address" does for the relay path.
    helpers.setNameFormula(anchorName(id), "=Model!$B$10");
    helpers.seed("Model!B10", [[42]]);
    const collector = new LocalCollector();

    const summary = await links.pushLinks("all", collector);

    expect(summary).toEqual({ pushed: 1, missing: 0, failed: 0, failures: [] });
    expect(collector.bundle().links[0]).toMatchObject({ id });
    const rows = await links.listWorkbookLinks();
    expect(rows[0]!.entry.label).toBe("Model!B4"); // the label is fixed at export
    expect(rows[0]!.source).toBe("ok");
  });
});

describe("Copy all, twice in a row", () => {
  it("both succeed and each advances the revision independently", async () => {
    helpers.seed("Model!C4", [[9]]);
    helpers.select("Model!B4");
    await links.exportSelection(ws, new LocalCollector());
    helpers.select("Model!C4");
    await links.exportSelection(ws, new LocalCollector());

    const first = new LocalCollector();
    const firstSummary = await links.pushLinks("all", first, { announce: ws });
    expect(firstSummary).toEqual({
      pushed: 2,
      missing: 0,
      failed: 0,
      failures: [],
    });
    expect(first.bundle().links).toHaveLength(2);

    const second = new LocalCollector();
    const secondSummary = await links.pushLinks("all", second, {
      announce: ws,
    });
    expect(secondSummary).toEqual({
      pushed: 2,
      missing: 0,
      failed: 0,
      failures: [],
    });
    for (const link of second.bundle().links) {
      const firstRev = first.bundle().links.find((l) => l.id === link.id)!.rev;
      expect(link.rev).toBeGreaterThan(firstRev);
    }
  });
});
