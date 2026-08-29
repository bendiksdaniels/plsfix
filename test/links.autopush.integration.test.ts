// Auto-push on edit (v2.1) against the strict fake host: an edit inside a
// linked range re-pushes that link once the typing stops, so "Update all" in
// PowerPoint always finds the current picture. The three-second window runs on
// a virtual clock here; the export and push flows themselves are covered in
// links.integration.test.ts.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import type * as WatchModule from "../src/excel/link-watch";
import { RelayError } from "../src/link/relay";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { virtualClock, type VirtualClock } from "./clock";
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
  watch = await import("../src/excel/link-watch");
  clock = virtualClock();
  notes = [];
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.select("Model!B4:F5");
});
afterEach(() => uninstallFakeHost());

// Reopening the file: the same workbook model, a brand new host runtime and
// module state, which is what makes the workbook flag worth keeping.
async function reopen(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ workbook });
  helpers = host.helpers;
  links = await import("../src/excel/links");
  watch = await import("../src/excel/link-watch");
}

// Long enough for a flush that decides to push nothing to have run.
async function settle(): Promise<void> {
  for (let turn = 0; turn < 5; turn += 1) {
    await new Promise((done) => setTimeout(done, 0));
  }
}

async function armed(): Promise<string> {
  const { id } = await links.exportSelection(ws, relay);
  await watch.setAutoPush(true, relay, note, { clock });
  return id;
}

describe("auto-push on edit", () => {
  it("pushes a link when a cell inside its anchor is edited", async () => {
    const id = await armed();
    expect(watch.autoPushEnabled()).toBe(true);

    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
    expect(notes).toEqual(["Pushed 1 link"]);
  });

  it("leaves the link alone when the edit lands outside it", async () => {
    const id = await armed();

    await helpers.fireChanged("Model", "H9");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();
    expect(relay.links.get(id)!.rev).toBe(1);
    expect(notes).toEqual([]);

    // The watcher is still live and still counting: rev 2, never 3.
    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
  });

  it("pushes once for two edits inside one window", async () => {
    const id = await armed();

    await helpers.fireChanged("Model", "B4");
    clock.advance(1000);
    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);

    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
    await settle();
    expect(relay.links.get(id)!.rev).toBe(2);
    expect(notes).toEqual(["Pushed 1 link"]);
  });

  // A chart has no range to intersect, so any change on its own sheet counts.
  it("re-pushes a chart when its worksheet changes", async () => {
    helpers.addChart("Model", {
      name: "Revenue bridge",
      width: 400,
      height: 200,
    });
    helpers.setActiveChart(workbook.charts[0]!);
    const { id } = await links.exportActiveChart(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireChanged("Model", "Z99");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
  });

  it("says what went wrong instead of throwing into the event", async () => {
    const id = await armed();
    relay.putLink = () => Promise.reject(new RelayError("server", "boom", 500));

    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(notes).toEqual(["Auto-push failed: Model!B4:F5: boom"]);
    });
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("stops watching, and clears the flag, when it is switched off", async () => {
    const id = await armed();
    await watch.setAutoPush(false, relay, note, { clock });

    expect(watch.autoPushEnabled()).toBe(false);
    expect(helpers.changeHandlerCount()).toBe(0);
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBe("");

    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("keeps the flag in the workbook and re-arms after it is reopened", async () => {
    const id = await armed();
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBe("1");
    expect(helpers.changeHandlerCount()).toBe(1);

    await reopen();
    expect(helpers.changeHandlerCount()).toBe(0);

    const second = virtualClock();
    expect(await watch.restoreAutoPush(relay, note, { clock: second })).toBe(
      true,
    );
    expect(helpers.changeHandlerCount()).toBe(1);

    await helpers.fireChanged("Model", "C5");
    second.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
  });

  it("stays off in a workbook that never asked for it", async () => {
    await links.exportSelection(ws, relay);
    expect(await watch.restoreAutoPush(relay, note, { clock })).toBe(false);
    expect(helpers.changeHandlerCount()).toBe(0);
  });
});
