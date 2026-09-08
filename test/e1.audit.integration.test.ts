// E1 audit: the Excel links side on paths the suite never walked - a chart
// link jumped to and removed, a structural edit under auto-push, an edit on a
// sheet with no links, a table that outgrew its cap after the export, and the
// relay unreachable at boot. Strict fake host throughout.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import type * as WatchModule from "../src/excel/link-watch";
import { anchorName, REGISTRY_SETTING } from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { virtualClock, type VirtualClock } from "./clock";
import { RelayError } from "../src/link/relay";
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

async function chartLink(): Promise<string> {
  helpers.addChart("Model", {
    name: "Revenue bridge",
    width: 400,
    height: 200,
  });
  helpers.setActiveChart(workbook.charts[0]!);
  return (await links.exportActiveChart(ws, relay)).id;
}

describe("a chart link's own flows", () => {
  it("jumps to a chart source by activating its sheet and the chart", async () => {
    const id = await chartLink();
    helpers.moveChart(anchorName(id), "Data");
    helpers.select("Model!A1");

    await links.goToSource(id);

    expect(workbook.activeChart?.name).toBe(anchorName(id));
    expect(helpers.sheet("Data").id).toBe(workbook.activeSheetId);
  });

  it("gives a chart its own name back when the link is removed", async () => {
    const id = await chartLink();
    expect(workbook.charts[0]!.name).toBe(anchorName(id));

    await links.removeLink(id, relay);

    expect(workbook.charts[0]!.name).toBe("Revenue bridge");
    expect(JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links).toEqual(
      [],
    );
    expect(relay.links.has(id)).toBe(false);
  });

  it("removes a chart link whose chart is gone without touching the workbook", async () => {
    const id = await chartLink();
    workbook.charts.length = 0;

    await links.removeLink(id, relay);

    expect(JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links).toEqual(
      [],
    );
  });
});

describe("auto-push and the sheets an edit did not touch", () => {
  it("leaves every link alone when the edited sheet holds none", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireChanged("Data", "A1");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    expect(relay.links.get(id)!.rev).toBe(1);
    expect(notes).toEqual([]);
  });

  it("asks the relay nothing at all in a workbook with no links", async () => {
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    expect(notes).toEqual([]);
    expect(relay.links.size).toBe(0);
  });
});

// Rows or columns moving under an anchor arrive with their own changeType and
// no address a link can be intersected against: every link on that sheet is
// re-pushed rather than matched.
describe("auto-push after a structural edit", () => {
  it("re-pushes a link the deleted rows never touched", async () => {
    helpers.seed("Model!H2", [[11, 12]]);
    helpers.select("Model!H2:I2");
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireStructuralChange("Model", "A20:A21", "RowDeleted");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
    expect(notes).toEqual(["Pushed 1 link"]);
  });

  it("re-pushes the shrunken range when rows inside the anchor are deleted", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });
    // Excel rewrites the hidden name when rows go: the anchor shrinks with it.
    helpers.setNameFormula(anchorName(id), "=Model!$B$4:$F$4");

    await helpers.fireStructuralChange("Model", "5:5", "RowDeleted");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(relay.links.get(id)!.rev).toBe(2);
    });
    // The list still reads "ok": a shrunken anchor is a live source, and only
    // a name Excel broke to #REF! reports missing.
    expect((await links.listWorkbookLinks())[0]!.source).toBe("ok");
  });

  it("leaves a sheet's links alone when the structural edit was elsewhere", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireStructuralChange("Data", "1:1", "RowInserted");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    expect(relay.links.get(id)!.rev).toBe(1);
    expect(notes).toEqual([]);
  });
});

// A modeller edits one sheet, tabs to another and edits there. The window is
// kept per worksheet, so the two fall due half a second apart: each sheet gets
// its own push, and neither link is pushed twice or missed.
describe("a burst of edits across two sheets", () => {
  async function twoLinks(): Promise<[string, string]> {
    const first = await links.exportSelection(ws, relay);
    helpers.seed("Data!A1", [[7, 8]]);
    helpers.select("Data!A1:B1");
    const second = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });
    return [first.id, second.id];
  }

  it("pushes the link on each sheet once, windows staggered", async () => {
    const [first, second] = await twoLinks();

    await helpers.fireChanged("Model", "C5");
    clock.advance(500);
    await helpers.fireChanged("Data", "A1");
    clock.advance(watch.AUTOPUSH_DELAY_MS - 500);
    await vi.waitFor(() => {
      expect(relay.links.get(first)!.rev).toBe(2);
    });
    expect(relay.links.get(second)!.rev).toBe(1);

    clock.advance(500);
    await vi.waitFor(() => {
      expect(relay.links.get(second)!.rev).toBe(2);
    });
    await settle();
    expect(relay.links.get(first)!.rev).toBe(2);
    expect(notes).toEqual(["Pushed 1 link", "Pushed 1 link"]);
  });

  it("pushes both in one window when the two edits fall due together", async () => {
    const [first, second] = await twoLinks();

    await helpers.fireChanged("Model", "C5");
    await helpers.fireChanged("Data", "A1");
    clock.advance(watch.AUTOPUSH_DELAY_MS);

    await vi.waitFor(() => {
      expect(notes).toEqual(["Pushed 2 links"]);
    });
    expect(relay.links.get(first)!.rev).toBe(2);
    expect(relay.links.get(second)!.rev).toBe(2);
  });
});

// The independence check: nothing in the Links tab may leave a workbook worse
// off than it found it when the relay cannot be reached at all, and no other
// tab may be taken down with it.
describe("with the relay unreachable", () => {
  function dead(): FakeRelay {
    const refuse = (): Promise<never> =>
      Promise.reject(new RelayError("network", "relay unreachable"));
    const out = new FakeRelay();
    out.putLink = refuse;
    out.getLink = refuse;
    out.deleteLink = refuse;
    out.status = refuse;
    out.touchLinks = refuse;
    out.fetchLinks = refuse;
    out.postInbox = refuse;
    out.listInbox = refuse;
    out.deleteInbox = refuse;
    return out;
  }

  it("refuses every export in one sentence and anchors nothing", async () => {
    const down = dead();
    helpers.addChart("Model", { name: "Revenue bridge" });
    helpers.setActiveChart(workbook.charts[0]!);

    for (const send of [
      () => links.exportSelection(ws, down),
      () => links.exportSelectionAsTable(ws, down),
      () => links.exportActiveChart(ws, down),
    ]) {
      await expect(send()).rejects.toThrow(/relay unreachable/);
    }
    helpers.select("Model!B4");
    await expect(links.exportSelectionAsText(ws, down)).rejects.toThrow(
      /relay unreachable/,
    );

    expect(workbook.names).toEqual([]);
    expect(workbook.charts[0]!.name).toBe("Revenue bridge");
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
  });

  it("counts a push as failed and keeps a removal's record to try again", async () => {
    const { id } = await links.exportSelection(ws, relay);
    const down = dead();

    expect(await links.pushLinks("all", down)).toMatchObject({
      pushed: 0,
      missing: 0,
      failed: 1,
    });
    await expect(links.removeLink(id, down)).rejects.toThrow(
      /nothing was removed, try again/,
    );
    await expect(
      (await import("../src/excel/link-touch")).touchWorkbookLinks(down),
    ).rejects.toThrow(/relay unreachable/);

    expect(
      JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links,
    ).toHaveLength(1);
    expect(workbook.names).toHaveLength(1);
  });

  it("leaves the list, the highlight and the other tabs working", async () => {
    await links.exportSelection(ws, relay);
    const excel = await import("../src/excel");

    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.source).toBe("ok");
    expect(await excel.toggleLinkHighlight()).toBe(true);
    expect(await excel.toggleLinkHighlight()).toBe(false);

    expect((await excel.listSheets()).map((one) => one.name)).toEqual([
      "Model",
      "Data",
    ]);
    expect((await excel.inspectSelection()).address).toBe("Model!B4:F5");
    expect(await excel.readWorkbookBrand()).toBeNull();
  });
});

describe("a table link that outgrew its cap", () => {
  it("reports the refusal per link instead of failing the whole push", async () => {
    helpers.seed("Model!B4", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!B4:C5");
    const { id } = await links.exportSelectionAsTable(ws, relay);
    // Rows inserted inside the block grow the hidden name past the table cap.
    helpers.setNameFormula(anchorName(id), "=Model!$B$4:$C$99");

    const summary = await links.pushLinks("all", relay);

    expect(summary.pushed).toBe(0);
    expect(summary.failed).toBe(1);
    expect(summary.failures[0]).toMatch(/60 rows/);
    expect(relay.links.get(id)!.rev).toBe(1);
  });
});

// The paint is a write like any other, and a sheet the modeller protected
// refuses it. The store must not go on claiming fills the sheet never took:
// the audit overlay asks it before painting, and the Links tab's own box is
// already back to clear by then.
describe("the highlight over a protected sheet", () => {
  it("owns nothing after the host refuses the paint", async () => {
    await links.exportSelection(ws, relay);
    const excel = await import("../src/excel");
    helpers.protectSheet("Model");

    await expect(excel.toggleLinkHighlight()).rejects.toThrow(/protected/);

    expect(excel.linkHighlightOn()).toBe(false);
    expect(helpers.fill("Model!B4").pattern).toBe("None");
    // Nothing for the next open to put back either.
    expect(helpers.setting("PLSFIX_LINK_HIGHLIGHT")).toBe("");
    // The other overlay is free to paint: nothing is standing in its way.
    helpers.select("Data!A1:B2");
    expect(await excel.toggleAuditOverlay()).toBe(true);
  });
});
