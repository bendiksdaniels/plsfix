// The Excel links adapter against a LocalCollector: an export always carries
// an inbox row (the existing publish() flow, unconditional in every
// transport), while a later Copy selected / Copy all only carries one when
// asked (`{ announce }`) - a relay-mode push never does. Local revisions
// live in their own space (src/link/local.ts), so a fresh collector still
// counts on from the registry's own revision.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  decodeInboxItem,
  decodePayload,
  REGISTRY_SETTING,
} from "../src/link/model";
import { LOCAL_REV_BASE, localWorkspace } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
import type { Workspace } from "../src/link/workspace";
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
let ws: Workspace;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  ws = await localWorkspace();
  links = await import("../src/excel/links");
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.select("Model!B4:F5");
});
afterEach(() => uninstallFakeHost());

describe("exporting and pushing against a LocalCollector", () => {
  it("carries one inbox row per export, one more per announced push, and none without announce", async () => {
    const collector = new LocalCollector();
    const result = await links.exportSelection(ws, collector);
    const bundle = collector.bundle();
    expect(bundle.links).toHaveLength(1);
    expect(bundle.links[0]).toMatchObject({
      id: result.id,
      rev: LOCAL_REV_BASE + 1,
    });
    expect(bundle.inbox).toHaveLength(1);

    const row = bundle.inbox[0]!;
    const item = decodeInboxItem(await open(ws.enc, ws.id, row.blob));
    expect(item).toMatchObject({ id: result.id, label: "Model!B4:F5" });

    const keys = await deriveLinkKeys(item.token);
    const payload = decodePayload(
      await open(keys.enc, result.id, bundle.links[0]!.blob),
    );
    expect(payload.kind).toBe("picture");

    // Change a cell, then Copy all with announce: a second recipient device
    // that never held the inbox row yet can still insert this link.
    helpers.seed("Model!B4", [[11]]);
    const second = new LocalCollector();
    const summary = await links.pushLinks("all", second, { announce: ws });
    expect(summary).toEqual({ pushed: 1, missing: 0, failed: 0, failures: [] });
    const secondBundle = second.bundle();
    expect(secondBundle.links).toHaveLength(1);
    expect(secondBundle.links[0]).toMatchObject({
      id: result.id,
      rev: LOCAL_REV_BASE + 2,
    });
    expect(secondBundle.inbox).toHaveLength(1);
    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
    expect(registry.links[0].rev).toBe(LOCAL_REV_BASE + 2);

    // A push with no announce - relay behaviour kept: the link travels, no
    // inbox row rides with it.
    const third = new LocalCollector();
    await links.pushLinks("all", third);
    const thirdBundle = third.bundle();
    expect(thirdBundle.links).toHaveLength(1);
    expect(thirdBundle.links[0]).toMatchObject({
      id: result.id,
      rev: LOCAL_REV_BASE + 3,
    });
    expect(thirdBundle.inbox).toEqual([]);
  });
});
