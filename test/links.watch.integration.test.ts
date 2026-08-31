// @vitest-environment jsdom
// The Links tab's own list following the workbook while the tab stays open:
// a source deleted or moved is not caught by the anchor until something
// re-reads the registry, and today only a push or opening the tab again does
// that. Driven through the real taskpane.html markup and tab switcher against
// the strict fake host, the way links.autopush.integration.test.ts drives
// auto-push; the debounce runs on a virtual clock here too.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { LIST_WATCH_DELAY_MS } from "../src/excel/link-list-watch";
import { anchorName } from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { installLinksTab, type LinksTabDeps } from "../src/pane/links-tab";
import type { Guard } from "../src/ui/guard";
import { installTabs } from "../src/ui/tabs";
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
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let ws: Workspace;
let clock: VirtualClock;
let errors: string[];

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
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  clock = virtualClock();
  errors = [];
  links = await import("../src/excel/links");
  helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  helpers.select("Model!B4:F5");
});
afterEach(() => uninstallFakeHost());

interface Pane {
  refresh(): Promise<void>;
  activate(tabId: string): void;
}

// The real pane markup and the real tab switcher, so a renamed id fails here
// instead of in Excel; only the relay is a fake, the way the adapter suites
// use it.
function install(): Pane {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const tabs = installTabs(document.getElementById("tab-bar")!);
  const guard: Guard = async (run) => {
    try {
      await run();
    } catch (error) {
      errors.push(error instanceof Error ? error.message : String(error));
    }
  };
  const deps: LinksTabDeps = {
    guard,
    toast: { show: () => undefined },
    relay,
    keyStore: memoryStore(),
    root: document,
    watch: { clock },
  };
  const handle = installLinksTab(deps);
  // The markup's own default tab is Tools; every test here is about a
  // modeller who already has the Links tab open.
  tabs.activate("tab-links");
  return { refresh: handle.refresh, activate: tabs.activate };
}

async function settle(): Promise<void> {
  for (let turn = 0; turn < 5; turn += 1) {
    await new Promise((done) => setTimeout(done, 0));
  }
}

function rowText(): string {
  return document.getElementById("workbook-links")?.textContent ?? "";
}

describe("the Links tab's list while it is open", () => {
  it("shows a broken source once the edit that broke it settles, without a push", async () => {
    const pane = install();
    await settle();
    const { id } = await links.exportSelection(ws, relay);
    await pane.refresh();
    expect(rowText()).not.toContain("Source missing");

    helpers.breakName(anchorName(id));
    await helpers.fireChanged("Model", "B4");
    clock.advance(LIST_WATCH_DELAY_MS);
    await settle();

    expect(rowText()).toContain("Source missing");
    expect(errors).toEqual([]);
  });

  it("refreshes once for two edits inside one window", async () => {
    const pane = install();
    await settle();
    await links.exportSelection(ws, relay);
    await pane.refresh();

    // What one refresh costs in round trips, measured directly: the window
    // below must cost exactly this much, not twice it.
    const beforeOne = helpers.syncCount();
    await pane.refresh();
    const oneRefresh = helpers.syncCount() - beforeOne;

    const before = helpers.syncCount();
    await helpers.fireChanged("Model", "B4");
    clock.advance(1000);
    await helpers.fireChanged("Model", "C5");
    clock.advance(LIST_WATCH_DELAY_MS);
    await settle();

    expect(helpers.syncCount() - before).toBe(oneRefresh);
  });

  it("does nothing while another tab is showing", async () => {
    const pane = install();
    await settle();
    await links.exportSelection(ws, relay);
    await pane.refresh();

    pane.activate("tab-brand");
    const before = helpers.syncCount();
    await helpers.fireChanged("Model", "B4");
    clock.advance(LIST_WATCH_DELAY_MS);
    await settle();
    expect(helpers.syncCount()).toBe(before);

    // Coming back to the Links tab still refreshes as it always has, on the
    // click alone - no edit or clock advance needed.
    pane.activate("tab-links");
    document.getElementById("tab-links")!.dispatchEvent(new Event("click"));
    await settle();
    expect(helpers.syncCount()).toBeGreaterThan(before);
  });
});
