// The Excel links adapter end to end against the strict fake host: a hidden
// name anchors the source, the registry lives in workbook settings, the picture
// is sealed onto the relay and announced to the inbox. Anchors, not addresses,
// are what a later push re-renders.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  anchorName,
  decodeInboxItem,
  decodePayload,
  REGISTRY_SETTING,
} from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { RelayError } from "../src/link/relay";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
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
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.select("Model!B4:F5");
});
afterEach(() => uninstallFakeHost());

async function payloadOf(id: string, token: string) {
  const keys = await deriveLinkKeys(token);
  const stored = relay.links.get(id)!;
  return decodePayload(await open(keys.enc, id, stored.blob));
}

describe("exportSelection", () => {
  it("anchors a hidden name, records the registry, pushes a sealed picture and posts to the inbox", async () => {
    const result = await links.exportSelection(ws, relay);
    expect(result.label).toBe("Model!B4:F5");
    const name = workbook.names.find((n) => n.name === anchorName(result.id))!;
    expect(name.visible).toBe(false);
    expect(name.formula).toBe("=Model!$B$4:$F$5");
    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
    expect(registry.links[0]).toMatchObject({
      id: result.id,
      kind: "range",
      anchor: anchorName(result.id),
      rev: 1,
    });
    const payload = await payloadOf(result.id, registry.links[0].token);
    expect(payload.src).toEqual({
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:F5",
      anchor: anchorName(result.id),
    });
    expect(payload.png).toBe(fakePng(5 * 64, 2 * 20));
    expect(payload.width).toBe(320);
    const inbox = await relay.listInbox(ws.id, ws.auth);
    expect(
      decodeInboxItem(await open(ws.enc, ws.id, inbox[0]!.blob)),
    ).toMatchObject({ id: result.id, label: "Model!B4:F5" });
  });
  it("needs a host that can render a picture", async () => {
    helpers.setSupported(() => false);
    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /Excel 2021/,
    );
    expect(workbook.names).toEqual([]);
  });
  it("binds the anchor before the upload and unbinds it when that fails", async () => {
    relay.putLink = () => Promise.reject(new RelayError("server", "boom", 500));
    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /export Model!B4:F5: boom/,
    );
    expect(workbook.names).toEqual([]);
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
  });
  it("refuses selections over the cap", async () => {
    helpers.select("Model!A:A");
    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /selected cells/,
    );
  });
});

describe("pushLinks", () => {
  it("renders through the anchor name, not the original address", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.setNameFormula(anchorName(id), "=Model!$B$10:$F$18");
    const summary = await links.pushLinks("all", relay);
    expect(summary).toEqual({
      pushed: 1,
      missing: 0,
      failed: 0,
      failures: [],
    });
    const token = JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0]
      .token;
    const payload = await payloadOf(id, token);
    expect(payload.src.ref).toBe("B10:F18");
    expect(payload.png).toBe(fakePng(5 * 64, 9 * 20));
    expect(relay.links.get(id)!.rev).toBe(2);
  });
  it("reports a broken anchor as missing and lists it as such", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.breakName(anchorName(id));
    expect(await links.pushLinks([id], relay)).toEqual({
      pushed: 0,
      missing: 1,
      failed: 0,
      failures: [],
    });
    expect((await links.listWorkbookLinks())[0]!.source).toBe("missing");
  });
});

describe("charts", () => {
  it("exports the active chart by renaming it to the anchor and finds it on another sheet later", async () => {
    helpers.addChart("Model", {
      name: "Revenue bridge",
      width: 400,
      height: 200,
    });
    helpers.setActiveChart(workbook.charts[0]!);
    const result = await links.exportActiveChart(ws, relay);
    expect(result.label).toBe("Model: Revenue bridge");
    expect(workbook.charts[0]!.name).toBe(anchorName(result.id));
    const token = JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0]
      .token;
    expect((await payloadOf(result.id, token)).png).toBe(fakePng(800, 400));
    helpers.moveChart(anchorName(result.id), "Data");
    expect(await links.pushLinks("all", relay)).toEqual({
      pushed: 1,
      missing: 0,
      failed: 0,
      failures: [],
    });
    expect((await payloadOf(result.id, token)).src.sheet).toBe("Data");
  });
  it("refuses to re-anchor a chart that is already linked", async () => {
    helpers.addChart("Model", { name: "Revenue bridge" });
    helpers.setActiveChart(workbook.charts[0]!);
    const first = await links.exportActiveChart(ws, relay);
    await expect(links.exportActiveChart(ws, relay)).rejects.toThrow(
      /already linked as Model: Revenue bridge; push it instead/,
    );
    expect(workbook.charts[0]!.name).toBe(anchorName(first.id));
  });
  it("gives a chart its name back when the upload fails", async () => {
    helpers.addChart("Model", { name: "Revenue bridge" });
    helpers.setActiveChart(workbook.charts[0]!);
    relay.putLink = () => Promise.reject(new RelayError("server", "boom", 500));
    await expect(links.exportActiveChart(ws, relay)).rejects.toThrow(/boom/);
    expect(workbook.charts[0]!.name).toBe("Revenue bridge");
  });
  it("needs a selected chart", async () => {
    helpers.setActiveChart(null);
    await expect(links.exportActiveChart(ws, relay)).rejects.toThrow(
      /Select a chart/,
    );
  });
});

describe("a registry that cannot be read", () => {
  it("is never overwritten", async () => {
    await links.exportSelection(ws, relay);
    const garbage = '{"v":9,"links":"nope"';
    helpers.setSetting(REGISTRY_SETTING, garbage);
    await expect(links.pushLinks("all", relay)).rejects.toThrow(
      /registry SMT_LINKS: unreadable, not overwriting/,
    );
    expect(helpers.setting(REGISTRY_SETTING)).toBe(garbage);
    await expect(links.listWorkbookLinks()).rejects.toThrow(/unreadable/);
  });
});

describe("removeLink and goToSource", () => {
  it("cleans the name, the registry and the relay", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await links.removeLink(id, relay);
    expect(
      workbook.names.find((n) => n.name === anchorName(id)),
    ).toBeUndefined();
    expect(JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links).toEqual(
      [],
    );
    expect(relay.links.has(id)).toBe(false);
  });
  it("selects the source range", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.select("Data!A1");
    await links.goToSource(id);
    expect(workbook.selectionAddress()).toBe("Model!B4:F5");
  });
});
