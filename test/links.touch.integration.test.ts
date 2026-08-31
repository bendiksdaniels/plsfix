// The boot-time TTL refresh against the strict Excel fake: every link the
// registry still holds is named to the relay with its own auth key, and a
// workbook without links asks the relay nothing at all.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { REGISTRY_SETTING } from "../src/link/model";
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
  helpers.select("Model!B4:C5");
});
afterEach(() => uninstallFakeHost());

describe("touchWorkbookLinks", () => {
  it("touches every link the registry holds, with its own auth", async () => {
    const { touchWorkbookLinks } = await import("../src/excel/link-touch");
    const { id } = await links.exportSelection(ws, relay);
    const touched = await touchWorkbookLinks(relay);
    expect(touched).toBe(1);
    expect(relay.touched).toEqual([
      expect.objectContaining({ id, auth: expect.any(String) }),
    ]);
    // The auth key is the link's own, never the workspace's.
    expect(relay.touched[0]!.auth).toBe(relay.links.get(id)!.auth);
  });

  it("asks the relay nothing for a workbook without links", async () => {
    const { touchWorkbookLinks } = await import("../src/excel/link-touch");
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
    expect(await touchWorkbookLinks(relay)).toBe(0);
    expect(relay.touched).toEqual([]);
  });
});
