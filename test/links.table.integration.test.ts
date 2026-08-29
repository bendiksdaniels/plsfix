// "Export as table" against the strict Excel fake: the anchor, the registry
// entry and the inbox note are a range export's, and the payload carries the
// text, the formats and the column widths of the selected cells. A cell wearing
// Excel's defaults carries nothing but its text, and a push re-renders through
// the anchor.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  anchorName,
  decodeInboxItem,
  decodePayload,
  REGISTRY_SETTING,
  type TablePayload,
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
  helpers.seed("Model!B4", [
    ["Revenue", 1000],
    ["Costs", -400],
  ]);
  helpers.setFont("Model!B4", { bold: true });
  helpers.setFont("Model!C5", { italic: true, color: "#FF0000", size: 9 });
  helpers.setFill("Model!C4", { color: "#EEEEEE" });
  helpers.setAlignment("Model!C4", "Right");
  helpers.sheet("Model").columnWidths.set(1, 96);
  helpers.sheet("Model").columnWidths.set(2, 48);
  helpers.select("Model!B4:C5");
});
afterEach(() => uninstallFakeHost());

function token(): string {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0].token;
}

async function tableOf(id: string): Promise<TablePayload> {
  const keys = await deriveLinkKeys(token());
  const stored = relay.links.get(id)!;
  const payload = decodePayload(await open(keys.enc, id, stored.blob));
  if (payload.kind !== "table") throw new Error(`${id}: not a table`);
  return payload;
}

describe("exportSelectionAsTable", () => {
  it("anchors the range, records a table entry and pushes the cells, formats and widths", async () => {
    const result = await links.exportSelectionAsTable(ws, relay);
    expect(result.label).toBe("Model!B4:C5 table");
    const name = workbook.names.find((n) => n.name === anchorName(result.id))!;
    expect(name.visible).toBe(false);
    expect(name.formula).toBe("=Model!$B$4:$C$5");
    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
    expect(registry.links[0]).toMatchObject({
      id: result.id,
      kind: "table",
      anchor: anchorName(result.id),
      rev: 1,
    });

    const payload = await tableOf(result.id);
    expect(payload).toMatchObject({ kind: "table", rows: 2, cols: 2 });
    expect(payload.widths).toEqual([96, 48]);
    expect(payload.src).toEqual({
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:C5",
      anchor: anchorName(result.id),
    });
    // Text as Excel displays it, and only what differs from the defaults.
    expect(payload.cells).toEqual([
      [
        { t: "Revenue", b: true },
        { t: "1000", f: "#EEEEEE", a: "r" },
      ],
      [{ t: "Costs" }, { t: "-400", i: true, c: "#FF0000", z: 9 }],
    ]);

    const inbox = await relay.listInbox(ws.id, ws.auth);
    expect(
      decodeInboxItem(await open(ws.enc, ws.id, inbox[0]!.blob)),
    ).toMatchObject({ kind: "table", label: "Model!B4:C5 table" });
  });

  it("re-renders through the anchor on a push, not through the address", async () => {
    const { id } = await links.exportSelectionAsTable(ws, relay);
    helpers.seed("Model!B10", [
      ["Profit", 600],
      ["Margin", "60%"],
    ]);
    helpers.setNameFormula(anchorName(id), "=Model!$B$10:$C$11");

    expect(await links.pushLinks("all", relay)).toEqual({
      pushed: 1,
      missing: 0,
      failed: 0,
      failures: [],
    });
    const payload = await tableOf(id);
    expect(payload.src.ref).toBe("B10:C11");
    expect(payload.cells.map((row) => row.map((cell) => cell.t))).toEqual([
      ["Profit", "600"],
      ["Margin", "60%"],
    ]);
    expect(relay.links.get(id)!.rev).toBe(2);
  });

  // The cap is a table's own, so it is refused before the hidden name is bound.
  it("refuses a selection past 60 rows or 20 columns and anchors nothing", async () => {
    helpers.select("Model!A1:U60");
    await expect(links.exportSelectionAsTable(ws, relay)).rejects.toThrow(
      "Tables go up to 60 rows and 20 columns; export a picture for more.",
    );
    expect(workbook.names).toEqual([]);
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
    expect(relay.links.size).toBe(0);
  });

  it("lists and removes a table link like any other", async () => {
    const { id } = await links.exportSelectionAsTable(ws, relay);
    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.entry.kind).toBe("table");
    expect(rows[0]!.source).toBe("ok");

    await links.goToSource(id);
    expect(workbook.selectionAddress()).toBe("Model!B4:C5");

    await links.removeLink(id, relay);
    expect(
      workbook.names.find((n) => n.name === anchorName(id)),
    ).toBeUndefined();
    expect(relay.links.has(id)).toBe(false);
  });
});
