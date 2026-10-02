// Critical, coordinator-assigned: localFloor (link.local-rev-floor.hunt2)
// jumps a fresh session's first local push to the clock, so the rev one
// below a tag's is not always exactly rev - 1. Pins that Revert still finds
// and restores the real previous payload in that case, and reverts once.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import type * as PptLinksModule from "../../src/ppt/links";
import type * as RevertModule from "../../src/ppt/revert";
import { localWorkspace } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import { LocalStore } from "../../src/link/local-store";
import { memoryPersistence } from "../../src/link/local-persist";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "../fakeppt";

enableStrictLoadSemantics();

let excelLinks: typeof LinksModule;
let excelHelpers: FakeHelpers;
let pptLinks: typeof PptLinksModule;
let revert: typeof RevertModule;
let presentation: FakePresentation;
let clockMs: number;

function tick(ms = 5000): void {
  clockMs += ms;
  vi.setSystemTime(clockMs);
}

// A fresh pane session against the SAME workbook: an upgrade reload, a
// reopened workbook's pane, or a Save As copy's own pane all reset
// link-record.ts's trustedThisSession the same way.
async function freshExcelSession(): Promise<typeof LinksModule> {
  vi.resetModules();
  return import("../../src/excel/links");
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  uninstallFakePpt();
  vi.useFakeTimers();
  clockMs = new Date("2026-01-01T00:00:00.000Z").getTime();
  vi.setSystemTime(clockMs);
  const excelHost = installFakeHost({ sheets: ["Model"] });
  excelHelpers = excelHost.helpers;
  excelLinks = await import("../../src/excel/links");
  excelHelpers.seed("Model!B4", [["v0"]]);
  excelHelpers.select("Model!B4");
  const pptHost = installFakePpt({ slides: 1 });
  presentation = pptHost.presentation;
  pptHost.helpers.selectSlide(presentation.slides[0]!.id);
  pptLinks = await import("../../src/ppt/links");
  revert = await import("../../src/ppt/revert");
});
afterEach(() => {
  uninstallFakeHost();
  uninstallFakePpt();
  vi.useRealTimers();
});

function firstShapeRev(rows: PptLinksModule.LinkRow[]): number {
  return rows[0]!.found.tag.rev;
}

describe("local mode: Revert across a session's floor jump", () => {
  it("restores the first payload and rev; a second Revert has nowhere to go", async () => {
    const ws = await localWorkspace();
    const store = await LocalStore.open(memoryPersistence());

    // Push #1, this session: export announces and pushes at a plain +1 rev.
    const collector1 = new LocalCollector();
    const { id } = await excelLinks.exportSelectionAsText(ws, collector1);
    await store.ingest(collector1.bundle(), new Set());
    const item = (await pptLinks.listInbox(ws, store)).find(
      (row) => row.id === id,
    )!;
    await pptLinks.insertFromInbox(item, ws, store);
    const firstRev = firstShapeRev(await pptLinks.listLinks(store));

    // Push #2, a FRESH session: trustedThisSession is empty, so localFloor
    // jumps this push's rev to the clock rather than firstRev + 1.
    tick();
    const afterUpgrade = await freshExcelSession();
    excelHelpers.seed("Model!B4", [["v1"]]);
    const collector2 = new LocalCollector();
    const pushSummary = await afterUpgrade.pushLinks([id], collector2);
    expect(pushSummary.failed).toBe(0);

    // Paste #2: the deck already holds this id, so it updates rather than
    // waiting in the Inbox - the everyday open, edit, Copy, paste, Revert.
    await store.ingest(collector2.bundle(), new Set([id]));
    const updateSummary = await pptLinks.updateLinks(
      await pptLinks.listLinks(store),
      store,
    );
    expect(updateSummary).toMatchObject({ updated: 1, failed: 0 });
    const jumpedRev = firstShapeRev(await pptLinks.listLinks(store));
    expect(jumpedRev).toBeGreaterThan(firstRev + 1);

    const summary = await revert.revertLinks(
      await pptLinks.listLinks(store),
      store,
    );
    expect(summary).toMatchObject({ reverted: 1, noPrevious: 0, failed: 0 });
    expect(firstShapeRev(await pptLinks.listLinks(store))).toBe(firstRev);

    // firstRev's own previous sits below previousRevOf's local-space floor.
    const second = await revert.revertLinks(
      await pptLinks.listLinks(store),
      store,
    );
    expect(second).toMatchObject({ reverted: 0, noPrevious: 1, failed: 0 });
  });

  it("still reverts exactly one step through a healthy, single-session chain", async () => {
    const ws = await localWorkspace();
    const store = await LocalStore.open(memoryPersistence());

    const collector1 = new LocalCollector();
    const { id } = await excelLinks.exportSelectionAsText(ws, collector1);
    await store.ingest(collector1.bundle(), new Set());
    const item = (await pptLinks.listInbox(ws, store)).find(
      (row) => row.id === id,
    )!;
    await pptLinks.insertFromInbox(item, ws, store);
    const firstRev = firstShapeRev(await pptLinks.listLinks(store));

    excelHelpers.seed("Model!B4", [["v1"]]);
    const collector2 = new LocalCollector();
    await excelLinks.pushLinks([id], collector2);
    await store.ingest(collector2.bundle(), new Set([id]));
    await pptLinks.updateLinks(await pptLinks.listLinks(store), store);
    expect(firstShapeRev(await pptLinks.listLinks(store))).toBe(firstRev + 1);

    const summary = await revert.revertLinks(
      await pptLinks.listLinks(store),
      store,
    );
    expect(summary).toMatchObject({ reverted: 1, noPrevious: 0, failed: 0 });
    expect(firstShapeRev(await pptLinks.listLinks(store))).toBe(firstRev);
  });
});
