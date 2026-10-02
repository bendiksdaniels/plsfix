// Pass-2 failures: a relay timeout or a garbled body, on export, push all
// and remove, plus a second press queued behind a first still uploading.
// Pin: the registry never records what did not land, and the next attempt
// works. Confirms two existing deliberate choices, fixes neither.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { RelayError } from "../../src/link/relay";
import { relayReason } from "../../src/link/relay-reason";
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

const TIMEOUT_ERROR = new RelayError(
  "timeout",
  "The link relay did not answer in time.",
);
const GARBLED_BODY_ERROR = new RelayError(
  "server",
  "relay PUT /api/links/x: bad response",
);

let links: typeof LinksModule;
let helpers: FakeHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
  helpers.seed("Model!B4", [[1]]);
  helpers.select("Model!B4");
});
afterEach(() => uninstallFakeHost());

describe("a relay timeout: the message is already plain, and nothing half-lands", () => {
  it("export: refuses cleanly, no anchor and no registry entry left behind, and a retry succeeds", async () => {
    relay.putLink = () => Promise.reject(TIMEOUT_ERROR);

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      "export Model!B4: The link relay did not answer in time.",
    );
    expect(await links.listWorkbookLinks()).toHaveLength(0);
    expect(helpers.setting("PLSFIX_LINKS")).toBeNull();

    relay.putLink = FakeRelay.prototype.putLink.bind(relay);
    const { id } = await links.exportSelection(ws, relay);
    expect(await links.listWorkbookLinks()).toHaveLength(1);
    expect(relay.links.get(id)).toBeDefined();
  });

  it("push all: the entry's rev and lastPushedAt stay exactly what export left, and a retry pushes clean", async () => {
    const { id } = await links.exportSelection(ws, relay);
    const before = (await links.listWorkbookLinks())[0]!.entry;
    relay.putLink = () => Promise.reject(TIMEOUT_ERROR);

    const failed = await links.pushLinks("all", relay);
    expect(failed).toEqual({
      pushed: 0,
      missing: 0,
      failed: 1,
      failures: ["Model!B4: The link relay did not answer in time."],
    });
    const stillBefore = (await links.listWorkbookLinks())[0]!.entry;
    expect(stillBefore).toEqual(before);

    relay.putLink = FakeRelay.prototype.putLink.bind(relay);
    const ok = await links.pushLinks("all", relay);
    expect(ok).toEqual({ pushed: 1, missing: 0, failed: 0, failures: [] });
    const after = (await links.listWorkbookLinks())[0]!.entry;
    // rev 1 came from the export's own first push; this retry is the second.
    expect(after.rev).toBe(2);
    expect(after.id).toBe(id);
  });

  it("remove: the entry stays in the workbook, and a retry removes it", async () => {
    const { id } = await links.exportSelection(ws, relay);
    relay.deleteLink = () => Promise.reject(TIMEOUT_ERROR);

    await expect(links.removeLink(id, relay)).rejects.toThrow(
      "remove Model!B4: The link relay did not answer in time.; nothing was removed, try again",
    );
    expect(await links.listWorkbookLinks()).toHaveLength(1);

    relay.deleteLink = FakeRelay.prototype.deleteLink.bind(relay);
    await links.removeLink(id, relay);
    expect(await links.listWorkbookLinks()).toHaveLength(0);
  });
});

describe("a garbled (wrong-shape) body: left as the relay said it, by design, but still nothing half-written", () => {
  it("relayReason leaves it alone, the same as a bare network error", () => {
    expect(relayReason(GARBLED_BODY_ERROR)).toBeUndefined();
  });

  it("export: refuses cleanly and a retry succeeds", async () => {
    relay.putLink = () => Promise.reject(GARBLED_BODY_ERROR);

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      "export Model!B4: relay PUT /api/links/x: bad response",
    );
    expect(await links.listWorkbookLinks()).toHaveLength(0);

    relay.putLink = FakeRelay.prototype.putLink.bind(relay);
    await links.exportSelection(ws, relay);
    expect(await links.listWorkbookLinks()).toHaveLength(1);
  });

  it("push all: fails the one link without touching its stored revision", async () => {
    await links.exportSelection(ws, relay);
    const before = (await links.listWorkbookLinks())[0]!.entry;
    relay.putLink = () => Promise.reject(GARBLED_BODY_ERROR);

    const summary = await links.pushLinks("all", relay);
    expect(summary.failed).toBe(1);
    expect((await links.listWorkbookLinks())[0]!.entry).toEqual(before);
  });
});

// Blocks the FIRST putLink call (never settling it), then every later call
// runs the real FakeRelay behaviour. Unlike FakeRelay's own holdNextPut,
// releasing here rejects the held call instead of letting it succeed - the
// scenario the brief asks for is a relay failure DURING a race, not just a
// slow one.
function blockFirstPutThenFail(relay: FakeRelay): {
  started: Promise<void>;
  failWith: (error: unknown) => void;
} {
  const original = relay.putLink.bind(relay);
  let calls = 0;
  let markStarted!: () => void;
  const started = new Promise<void>((resolve) => {
    markStarted = resolve;
  });
  let reject!: (error: unknown) => void;
  const gate = new Promise<never>((_resolve, doReject) => {
    reject = doReject;
  });
  relay.putLink = (id: string, auth: string, blob: Uint8Array) => {
    calls += 1;
    if (calls === 1) {
      markStarted();
      return gate as Promise<{ rev: number }>;
    }
    return original(id, auth, blob);
  };
  return { started, failWith: reject };
}

describe("a second press while the first is still mid-upload", () => {
  it("queues behind the first push all instead of interleaving, and both land correctly once the failed one is retried", async () => {
    helpers.seed("Model!C4", [[2]]);
    helpers.select("Model!C4");
    await links.exportSelection(ws, relay);
    helpers.select("Model!B4");
    await links.exportSelection(ws, relay);
    expect(await links.listWorkbookLinks()).toHaveLength(2);

    const gate = blockFirstPutThenFail(relay);
    const first = links.pushLinks("all", relay);
    await gate.started;

    // The second press starts while the first is still parked on its first
    // link's put: exclusive() must make it wait for the first to finish
    // pushing both links and writing the registry back, never read a
    // registry the first has not written back yet.
    const second = links.pushLinks("all", relay);

    gate.failWith(new RelayError("server", "relay PUT /api/links/x: 503", 503));
    const firstSummary = await first;
    const secondSummary = await second;

    // The first link in registry order (Model!C4, exported first) is the one
    // the gate held, so it is the one that failed; the second link in that
    // same batch (Model!B4) went through the real relay untouched.
    expect(firstSummary).toEqual({
      pushed: 1,
      missing: 0,
      failed: 1,
      failures: ["Model!C4: The relay had a problem. Try again in a minute."],
    });
    // The second press, entirely after the first wrote back, sees both
    // links and pushes both clean - including a fresh attempt at the one
    // the first press's gate had failed.
    expect(secondSummary).toEqual({
      pushed: 2,
      missing: 0,
      failed: 0,
      failures: [],
    });
    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(2);
    for (const row of rows) expect(row.source).toBe("ok");

    // A third pass proves the registry the two overlapping presses left
    // behind is coherent, not stuck on a half-applied write from either one.
    const third = await links.pushLinks("all", relay);
    expect(third).toEqual({ pushed: 2, missing: 0, failed: 0, failures: [] });
  });
});
