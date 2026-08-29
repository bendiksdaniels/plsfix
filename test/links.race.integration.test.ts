// The SMT_LINKS registry under two flows at once. Every link flow reads the
// registry, does network I/O and writes the whole thing back, and the auto-push
// window runs from a timer rather than from a click, so the Links tab's own
// busy state cannot keep the two apart. The relay here holds an upload open the
// way a slow network does, which is the window a modeller hits by clicking
// Export while the debounce that started when they stopped typing falls due.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import type * as LockModule from "../src/excel/link-lock";
import type * as WatchModule from "../src/excel/link-watch";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { virtualClock, type VirtualClock } from "./clock";
import { FakeRelay, type HeldPut } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let lock: typeof LockModule;
let watch: typeof WatchModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let ws: Workspace;
let clock: VirtualClock;
let notes: string[];

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

function note(message: string): void {
  notes.push(message);
}

// Long enough for anything that was going to run without waiting on the relay.
async function settle(): Promise<void> {
  for (let turn = 0; turn < 5; turn += 1) {
    await new Promise((done) => setTimeout(done, 0));
  }
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
  lock = await import("../src/excel/link-lock");
  watch = await import("../src/excel/link-watch");
  clock = virtualClock();
  notes = [];
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.seed("Model!H2", [[11, 12]]);
});
afterEach(() => uninstallFakeHost());

// An auto-push stopped inside its upload, with the edit that started it already
// counted: the registry has been read and is about to be written back.
async function autoPushInFlight(): Promise<HeldPut> {
  await watch.setAutoPush(true, relay, note, { clock });
  const held = relay.holdNextPut();
  await helpers.fireChanged("Model", "C5");
  clock.advance(watch.AUTOPUSH_DELAY_MS);
  await held.started;
  return held;
}

describe("the link registry under two flows at once", () => {
  it("keeps the exported link when an auto-push was mid-upload", async () => {
    helpers.select("Model!B4:F5");
    const first = await links.exportSelection(ws, relay);
    const held = await autoPushInFlight();
    expect(lock.linkQueueStage()).toBe("auto-push");

    // The modeller clicks Export while the upload is still open.
    helpers.select("Model!H2:I3");
    const exporting = links.exportSelection(ws, relay);
    await settle();
    // Queued behind the push, so it has not read the registry yet.
    expect((await links.listWorkbookLinks()).length).toBe(1);

    held.release();
    const second = await exporting;
    await vi.waitFor(() => {
      expect(notes).toEqual(["Pushed 1 link"]);
    });

    const rows = await links.listWorkbookLinks();
    expect(rows.map((row) => row.entry.id).sort()).toEqual(
      [first.id, second.id].sort(),
    );
    expect(rows.every((row) => row.source === "ok")).toBe(true);
    // The new link is on the relay and reachable from the registry that names
    // it: the token that revokes it did not go missing with the entry.
    expect(relay.links.has(second.id)).toBe(true);
    expect(lock.linkQueueStage()).toBeNull();
  });

  it("keeps a removed link removed when an auto-push was mid-upload", async () => {
    helpers.select("Model!B4:F5");
    const first = await links.exportSelection(ws, relay);
    helpers.select("Model!H2:I3");
    const second = await links.exportSelection(ws, relay);

    const held = await autoPushInFlight();
    const removing = links.removeLink(second.id, relay);
    await settle();

    held.release();
    await removing;
    await vi.waitFor(() => {
      expect(notes).toEqual(["Pushed 1 link"]);
    });

    // The entry stays gone, and so does its relay copy: an auto-push that
    // finished afterwards must not write the registry it read back over it.
    const rows = await links.listWorkbookLinks();
    expect(rows.map((row) => row.entry.id)).toEqual([first.id]);
    expect(relay.links.has(second.id)).toBe(false);
    expect(workbook.names.map((name) => name.name)).not.toContain(
      `SMT_LINK_${second.id.slice(0, 8)}`,
    );
  });
});
