// "Export as text" against the strict Excel fake: one cell's displayed text
// travels as the text payload under a "Sheet!A1 text" label, and two cells, an
// empty cell or more than 500 characters are refused before anything is
// anchored. A push re-renders the text through the same hidden name.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  anchorName,
  decodeInboxItem,
  decodePayload,
  REGISTRY_SETTING,
  type Payload,
  type RegistryEntry,
} from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";

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
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../src/excel/links");
  helpers.seed("Model!B4", [["EUR 15.7m", 1000]]);
  helpers.select("Model!B4");
});
afterEach(() => uninstallFakeHost());

function registryEntry(): RegistryEntry {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0];
}

function names(): string[] {
  return workbook.names.map((name) => name.name);
}

async function payloadOf(id: string, token: string): Promise<Payload> {
  const keys = await deriveLinkKeys(token);
  const result = await relay.getLink(id, keys.auth);
  if (result === "unchanged") throw new Error("no blob");
  return decodePayload(await open(keys.enc, id, result.blob));
}

describe("Export as text", () => {
  it("sends the displayed text of one cell under a text label", async () => {
    const { label } = await links.exportSelectionAsText(ws, relay);
    expect(label).toBe("Model!B4 text");
    const entry = registryEntry();
    expect(entry.kind).toBe("text");
    expect(names()).toContain(anchorName(entry.id));

    const payload = await payloadOf(entry.id, entry.token);
    expect(payload.kind).toBe("text");
    expect(payload.kind === "text" && payload.text).toBe("EUR 15.7m");
    expect(payload.src).toEqual({
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4",
      anchor: anchorName(entry.id),
    });

    const inbox = await relay.listInbox(ws.id, ws.auth);
    expect(
      decodeInboxItem(await open(ws.enc, ws.id, inbox[0]!.blob)),
    ).toMatchObject({ kind: "text", label: "Model!B4 text" });
  });

  it("refuses two cells", async () => {
    helpers.select("Model!B4:C4");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow(
      "Select one cell",
    );
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
    expect(names()).toEqual([]);
  });

  it("refuses an empty cell", async () => {
    helpers.select("Model!B9");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow(
      "empty",
    );
    expect(names()).toEqual([]);
  });

  it("refuses more than 500 characters before anchoring", async () => {
    helpers.seed("Model!B6", [["x".repeat(501)]]);
    helpers.select("Model!B6");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow(
      "500 characters",
    );
    expect(names()).toEqual([]);
    expect(relay.links.size).toBe(0);
  });

  it("re-renders the text through the anchor on push", async () => {
    await links.exportSelectionAsText(ws, relay);
    const entry = registryEntry();
    helpers.seed("Model!B4", [["EUR 16.1m"]]);

    expect(await links.pushLinks([entry.id], relay)).toMatchObject({
      pushed: 1,
      failed: 0,
    });
    const payload = await payloadOf(entry.id, entry.token);
    expect(payload.kind === "text" && payload.text).toBe("EUR 16.1m");
    expect(relay.links.get(entry.id)!.rev).toBe(2);
  });

  it("lists and removes a text link like any other", async () => {
    const { id } = await links.exportSelectionAsText(ws, relay);
    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.entry.kind).toBe("text");
    expect(rows[0]!.source).toBe("ok");

    await links.removeLink(id, relay);
    expect(names()).toEqual([]);
    expect(relay.links.has(id)).toBe(false);
  });
});
