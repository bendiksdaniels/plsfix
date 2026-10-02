// Attacks src/excel/link-record.ts directly: the payload a table push seals
// (the header flag), the project an announce carries, and publish()'s
// rollback - both the ordinary case and the double failure where rollback's
// own sync also refuses, which must never bury the original error.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { deriveLinkKeys, open } from "../../src/link/crypto";
import {
  decodeInboxItem,
  decodePayload,
  REGISTRY_SETTING,
} from "../../src/link/model";
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
  type FakeWorkbook,
} from "../fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
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

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
});
afterEach(() => uninstallFakeHost());

async function inboxItem(): Promise<Record<string, unknown>> {
  const rows = await relay.listInbox(ws.id, ws.auth);
  const item = decodeInboxItem(await open(ws.enc, ws.id, rows[0]!.blob));
  return item as unknown as Record<string, unknown>;
}

function registryToken(id: string): string {
  const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING))) as {
    links: { id: string; token: string }[];
  };
  return registry.links.find((link) => link.id === id)!.token;
}

async function tablePayload(id: string) {
  const keys = await deriveLinkKeys(registryToken(id));
  return decodePayload(await open(keys.enc, id, relay.links.get(id)!.blob));
}

describe("a table export's header flag", () => {
  it("carries h:true when every non-empty cell of the first row came out bold", async () => {
    helpers.seed("Model!B4", [
      ["Revenue", "Costs"],
      [100, -40],
    ]);
    helpers.setFont("Model!B4", { bold: true });
    helpers.setFont("Model!C4", { bold: true });
    helpers.select("Model!B4:C5");

    const result = await links.exportSelectionAsTable(ws, relay);
    const payload = await tablePayload(result.id);
    if (payload.kind !== "table") throw new Error("not a table");
    expect(payload.h).toBe(true);
  });

  it("omits h entirely for a plain, non-bold table", async () => {
    helpers.seed("Model!B4", [
      ["Revenue", "Costs"],
      [100, -40],
    ]);
    helpers.select("Model!B4:C5");

    const result = await links.exportSelectionAsTable(ws, relay);
    const payload = await tablePayload(result.id);
    if (payload.kind !== "table") throw new Error("not a table");
    expect(payload.h).toBeUndefined();
  });

  it("omits h when only some of the first row's non-empty cells are bold", async () => {
    helpers.seed("Model!B4", [
      ["Revenue", "Costs"],
      [100, -40],
    ]);
    helpers.setFont("Model!B4", { bold: true });
    helpers.select("Model!B4:C5");

    const result = await links.exportSelectionAsTable(ws, relay);
    const payload = await tablePayload(result.id);
    if (payload.kind !== "table") throw new Error("not a table");
    expect(payload.h).toBeUndefined();
  });
});

describe("announce and the active project", () => {
  it("carries the workbook's active project on the inbox note", async () => {
    helpers.seed("Model!B4", [[1, 2]]);
    helpers.select("Model!B4:C4");
    helpers.setSetting(
      REGISTRY_SETTING,
      JSON.stringify({ v: 1, links: [], activeProject: "Amasty" }),
    );

    await links.exportSelection(ws, relay);

    expect((await inboxItem()).project).toBe("Amasty");
  });

  it("carries no project field at all when nothing is active", async () => {
    helpers.seed("Model!B4", [[1, 2]]);
    helpers.select("Model!B4:C4");

    await links.exportSelection(ws, relay);

    expect("project" in (await inboxItem())).toBe(false);
  });
});

describe("publish's rollback", () => {
  beforeEach(() => {
    helpers.seed("Model!B4", [[1, 2]]);
    helpers.select("Model!B4:C4");
  });

  it("undoes the anchor and the registry write when the announce fails", async () => {
    relay.postInbox = () => Promise.reject(new Error("inbox is full"));

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /export Model!B4:C4: inbox is full/,
    );

    expect(workbook.names).toEqual([]);
    // A fresh workbook never had a registry setting; rollback writes back
    // exactly the registry it read before this export (link-anchors.ts's
    // emptyRegistry(), an empty-but-present object, never a bare null), so
    // the invariant that matters is the link count, not the raw encoding.
    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
    expect(registry.links).toEqual([]);
  });

  // The stronger version of the test above: rolling back onto an EMPTY
  // workbook cannot tell "restored the prior state" apart from "never wrote
  // anything at all". Seeding one real link first closes that gap: a
  // rollback that clobbered link.registry instead of restoring it would
  // drop this pre-existing entry too.
  it("restores exactly the prior registry, keeping an earlier link untouched", async () => {
    helpers.seed("Model!B10", [[9]]);
    helpers.select("Model!B10");
    const first = await links.exportSelection(ws, relay);

    helpers.select("Model!B4:C4");
    relay.postInbox = () => Promise.reject(new Error("inbox is full"));
    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /export Model!B4:C4: inbox is full/,
    );

    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING))) as {
      links: { id: string }[];
    };
    expect(registry.links.map((entry) => entry.id)).toEqual([first.id]);
  });

  // Measured empirically (helpers.syncCount() after the ordinary single
  // failure above): this scenario runs exactly 6 Excel syncs, the sixth
  // being rollback's own commit of the undo. Failing that one too must
  // still surface the ORIGINAL announce error, per the "best effort"
  // comment on rollback() - never a rollback error masking it, and never a
  // thrown error with no message at all.
  it("still throws the original announce error when rollback's own sync also refuses", async () => {
    // An earlier, untouched link first: release() deletes a name the moment
    // it runs (synchronously, before rollback's own sync even starts), so
    // checking names against zero would prove nothing about which sync
    // failed. A survivor from a DIFFERENT export is the proof instead -
    // ending at exactly one shows this export's own anchor was still
    // cleaned up despite the sync failing, and the earlier link's was not
    // touched by it.
    helpers.seed("Model!B10", [[9]]);
    helpers.select("Model!B10");
    await links.exportSelection(ws, relay);

    helpers.select("Model!B4:C4");
    relay.postInbox = () => Promise.reject(new Error("inbox is full"));
    helpers.failNextSync(new Error("disk full"), 5);

    const failure = (await links
      .exportSelection(ws, relay)
      .then(() => null)
      .catch((error: unknown) => error)) as Error;

    expect(failure).toBeInstanceOf(Error);
    expect(failure.message).toMatch(/export Model!B4:C4: inbox is full/);
    expect(failure.message).not.toMatch(/disk full/);
    expect(workbook.names).toHaveLength(1);
  });
});
