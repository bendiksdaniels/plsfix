// E1 audit: the Excel links side on paths the suite never walked - a chart
// link jumped to and removed, a structural edit under auto-push, an edit on a
// sheet with no links, a table that outgrew its cap after the export, and the
// relay unreachable at boot. Strict fake host throughout.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import type * as WatchModule from "../src/excel/link-watch";
import { deriveLinkKeys, open } from "../src/link/crypto";
import { anchorName, decodePayload, REGISTRY_SETTING } from "../src/link/model";
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

async function payloadSource(id: string): Promise<unknown> {
  const entry = JSON.parse(
    String(helpers.setting(REGISTRY_SETTING)),
  ).links.find((link: { id: string }) => link.id === id);
  const keys = await deriveLinkKeys(entry.token);
  const stored = relay.links.get(id)!;
  return decodePayload(await open(keys.enc, id, stored.blob)).src;
}

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
    // Since slice B the overlay owns only fills it striped, so the block gets
    // a consistent formula row to stripe.
    helpers.seed("Data!A1", [
      [
        { formula: "=C1*2", r1c1: "=RC[2]*2", value: 2 },
        { formula: "=D1*2", r1c1: "=RC[2]*2", value: 4 },
      ],
    ]);
    helpers.select("Data!A1:B2");
    expect(await excel.toggleAuditOverlay()).toBe(true);
  });
});

describe("what the host refuses before anything is anchored", () => {
  it("refuses an export outright when the registry cannot be read", async () => {
    helpers.failNextSync();
    await expect(links.exportSelection(ws, relay)).rejects.toThrow();
    expect(workbook.names).toEqual([]);
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
    expect(relay.links.size).toBe(0);

    helpers.failNextSync();
    await expect(links.listWorkbookLinks()).rejects.toThrow();
  });

  it("says which chart the pick named when it is not on this sheet", async () => {
    helpers.addChart("Model", { name: "Revenue bridge" });
    helpers.addChart("Data", { name: "Segment pie" });
    helpers.setActiveChart(null);

    await expect(
      links.exportActiveChart(ws, relay, "Segment pie"),
    ).rejects.toThrow("No chart named Segment pie on this sheet.");
    expect(workbook.names).toEqual([]);
  });

  it("reports a source that is gone when the jump is asked for", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.breakName(anchorName(id));

    await expect(links.goToSource(id)).rejects.toThrow(
      "go to source Model!B4:F5: source missing",
    );
  });

  it("names the flow and the id for a link this workbook never had", async () => {
    const stranger = "c".repeat(32);
    await expect(links.removeLink(stranger, relay)).rejects.toThrow(
      "remove " + stranger + ": not in this workbook",
    );
    await expect(links.pushLinks([stranger], relay)).rejects.toThrow(
      "push " + stranger + ": not in this workbook",
    );
  });
});

// The workbook name travels in every payload and is what the deck's source
// filter groups by. It comes from the Common API, not Excel's: a host that
// carries no document surface, and an unsaved workbook, both read as "".
describe("the name of the workbook a link came from", () => {
  it("is empty when the host offers no file properties", async () => {
    const office = Office as unknown as { context: { document: unknown } };
    const document = office.context.document;
    office.context.document = { addHandlerAsync: () => undefined };
    try {
      const { id } = await links.exportSelection(ws, relay);
      expect(await payloadSource(id)).toMatchObject({ workbook: "" });
    } finally {
      office.context.document = document;
    }
  });

  it("is empty for a workbook that was never saved", async () => {
    workbook.fileUrl = "";
    const { id } = await links.exportSelection(ws, relay);
    expect(await payloadSource(id)).toMatchObject({
      workbook: "",
      sheet: "Model",
      ref: "B4:F5",
    });
  });
});

// The anchor is a hidden name, and the Name Manager can delete one. That is
// not the same as a name Excel broke to #REF!, and it has to read the same
// way: the source is gone, and Remove still cleans up after it.
describe("a hidden name the modeller deleted", () => {
  it("lists the link as missing and still removes it", async () => {
    const { id } = await links.exportSelection(ws, relay);
    workbook.names.length = 0;

    expect((await links.listWorkbookLinks())[0]!.source).toBe("missing");
    expect(await links.pushLinks("all", relay)).toMatchObject({
      pushed: 0,
      missing: 1,
      failed: 0,
    });
    await expect(links.goToSource(id)).rejects.toThrow(/source missing/);

    await links.removeLink(id, relay);
    expect(JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links).toEqual(
      [],
    );
    expect(relay.links.has(id)).toBe(false);
  });
});

// Nothing may throw out of the event: the window runs from a timer, where a
// rejection has no caller to reach.
describe("what an auto-push window does with a workbook it cannot read", () => {
  it("reports the reason through notify instead of throwing", async () => {
    await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });
    helpers.setSetting(REGISTRY_SETTING, '{"v":9,"links":"nope"');

    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await vi.waitFor(() => {
      expect(notes).toEqual([
        "Auto-push failed: registry PLSFIX_LINKS: unreadable, not overwriting",
      ]);
    });
  });

  // Off means off, even for a window that had already fallen due and was
  // queued behind a push that was still uploading.
  it("drops a window queued behind a slow push once it is switched off", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    const held = relay.holdNextPut();
    const pushing = links.pushLinks("all", relay);
    await held.started;
    await helpers.fireChanged("Model", "C5");
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    await watch.setAutoPush(false, relay, note, { clock });
    held.release();
    await pushing;
    await settle();

    expect(relay.links.get(id)!.rev).toBe(2);
    expect(notes).toEqual([]);
  });

  it("skips a link whose source went away inside the window", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });

    await helpers.fireChanged("Model", "C5");
    helpers.breakName(anchorName(id));
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    expect(relay.links.get(id)!.rev).toBe(1);
    expect(notes).toEqual([]);
  });
});

// The selections a modeller really makes, at the edges the caps name.
describe("edge selections", () => {
  it("refuses a merged block as text and takes it as a picture", async () => {
    helpers.merge("Model!B4:C5");
    helpers.select("Model!B4:C5");

    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow(
      "Select one cell for a text link (merged cells: export as a picture).",
    );
    expect(workbook.names).toEqual([]);

    const picture = await links.exportSelection(ws, relay);
    expect(picture.label).toBe("Model!B4:C5");
  });

  it("carries unicode and a five-hundredth character through a text link", async () => {
    const text = `€ ${"ā".repeat(200)} ${"→".repeat(200)}`;
    helpers.seed("Model!H2", [[text]]);
    helpers.select("Model!H2");

    const { id } = await links.exportSelectionAsText(ws, relay);
    const entry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)))
      .links[0];
    const keys = await deriveLinkKeys(entry.token);
    const stored = relay.links.get(id)!;
    const payload = decodePayload(await open(keys.enc, id, stored.blob));
    expect(payload.kind === "text" && payload.text).toBe(text);
  });

  it("takes a single cell as a picture and refuses a whole column", async () => {
    helpers.select("Model!B4");
    expect((await links.exportSelection(ws, relay)).label).toBe("Model!B4");

    helpers.select("Model!A:A");
    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      "Export supports up to 5,000 selected cells at once.",
    );
    expect(workbook.names).toHaveLength(1);
  });

  it("keeps a link on a hidden sheet listed but refuses the jump", async () => {
    helpers.addSheet("Calc");
    helpers.seed("Calc!A1", [[1, 2]]);
    helpers.select("Calc!A1:B1");
    const { id } = await links.exportSelection(ws, relay);
    helpers.sheet("Calc").visibility = "Hidden";

    const rows = await links.listWorkbookLinks();
    expect(rows[0]!.source).toBe("ok");
    expect(await links.pushLinks("all", relay)).toMatchObject({ pushed: 1 });
    await expect(links.goToSource(id)).rejects.toThrow(/is hidden/);
  });
});

// A model made on Microsoft 365 and opened in Excel 2019 (ExcelApi 1.8) still
// lists its links and still offers Push. Range.getImage and the chart image
// surface are both 1.9, so the picture kinds have to say what the export says
// rather than reaching for a method the host does not carry.
describe("a push from a host below the picture floor", () => {
  it("refuses a picture link in the pane's own sentence", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.setSupported(() => false);

    const summary = await links.pushLinks("all", relay);

    expect(summary).toMatchObject({ pushed: 0, missing: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(/Excel 2021 \/ Microsoft 365 required/);
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("refuses a chart link the same way", async () => {
    const id = await chartLink();
    helpers.setSupported(() => false);

    const summary = await links.pushLinks("all", relay);

    expect(summary.failed).toBe(1);
    expect(summary.failures[0]).toMatch(/Excel 2021 \/ Microsoft 365 required/);
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("still pushes a text link, which needs nothing above 1.1", async () => {
    helpers.seed("Model!H2", [["Revenue"]]);
    helpers.select("Model!H2");
    const { id } = await links.exportSelectionAsText(ws, relay);
    helpers.setSupported(() => false);

    expect(await links.pushLinks("all", relay)).toMatchObject({
      pushed: 1,
      failed: 0,
    });
    expect(relay.links.get(id)!.rev).toBe(2);
  });
});

// The flag is written only once the workbook really is watching: a failed
// registration that still left "1" behind would re-arm on every later open of a
// file whose handler never took.
describe("arming auto-push against a host that refuses", () => {
  it("stays off, keeps the flag clear and can be armed again", async () => {
    await links.exportSelection(ws, relay);
    helpers.failNextSync();

    await expect(
      watch.setAutoPush(true, relay, note, { clock }),
    ).rejects.toThrow();
    expect(watch.autoPushEnabled()).toBe(false);
    expect(helpers.changeHandlerCount()).toBe(0);
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBeNull();

    await watch.setAutoPush(true, relay, note, { clock });
    expect(watch.autoPushEnabled()).toBe(true);
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBe("1");
  });

  it("writes the flag off for a workbook that was never watching", async () => {
    await watch.setAutoPush(false, relay, note, { clock });
    expect(watch.autoPushEnabled()).toBe(false);
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBe("");
  });

  // Two clicks in one turn: the toggles are serialized, so the handler can
  // never be registered twice and the last click is what the workbook keeps.
  it("ends where the last of two overlapping toggles asked", async () => {
    const on = watch.setAutoPush(true, relay, note, { clock });
    const off = watch.setAutoPush(false, relay, note, { clock });
    await Promise.all([on, off]);

    expect(watch.autoPushEnabled()).toBe(false);
    expect(helpers.changeHandlerCount()).toBe(0);
    expect(helpers.setting(watch.AUTOPUSH_SETTING)).toBe("");
  });

  it("drops a window that was already waiting when it is switched off", async () => {
    const { id } = await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, note, { clock });
    await helpers.fireChanged("Model", "C5");

    await watch.setAutoPush(false, relay, note, { clock });
    clock.advance(watch.AUTOPUSH_DELAY_MS);
    await settle();

    expect(relay.links.get(id)!.rev).toBe(1);
    expect(notes).toEqual([]);
  });
});

// The list watcher has no on/off switch: a host that refuses the registration
// leaves the caller without a live view and nothing else.
describe("the list watcher against a host that refuses", () => {
  it("swallows the refusal and never calls back", async () => {
    const { watchWorksheetEdits, LIST_WATCH_DELAY_MS } =
      await import("../src/excel/link-list-watch");
    const settled: number[] = [];
    helpers.failNextSync();

    watchWorksheetEdits(() => settled.push(1), { clock });
    await settle();
    await helpers.fireChanged("Model", "C5");
    clock.advance(LIST_WATCH_DELAY_MS);
    await settle();

    expect(settled).toEqual([]);
    expect(helpers.changeHandlerCount()).toBe(0);
  });
});

// The inbox note is posted after the registry is already written: a relay that
// takes the picture and refuses the note has to leave the workbook as it was.
describe("an inbox note the relay refuses", () => {
  it("takes the whole export back", async () => {
    relay.postInbox = () =>
      Promise.reject(new RelayError("server", "inbox full", 507));

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      /export Model!B4:F5: inbox full/,
    );
    expect(workbook.names).toEqual([]);
    expect(JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links).toEqual(
      [],
    );
  });

  it("leaves an earlier link untouched when a later export is taken back", async () => {
    const first = await links.exportSelection(ws, relay);
    relay.postInbox = () =>
      Promise.reject(new RelayError("server", "inbox full", 507));
    helpers.seed("Data!A1", [[7, 8]]);
    helpers.select("Data!A1:B1");

    await expect(links.exportSelection(ws, relay)).rejects.toThrow();

    const rows = await links.listWorkbookLinks();
    expect(rows.map((row) => row.entry.id)).toEqual([first.id]);
    expect(workbook.names.map((name) => name.name)).toEqual([
      anchorName(first.id),
    ]);
  });
});

// A defined name is workbook state, not sheet state, and a picture is a read:
// protecting the sheet a model lives on is normal, and it must not stop the
// modeller sending a picture of it.
describe("a source on a protected sheet", () => {
  it("exports, lists, pushes and removes like any other", async () => {
    helpers.protectSheet("Model");

    const { id } = await links.exportSelection(ws, relay);
    expect(workbook.names.map((name) => name.name)).toEqual([anchorName(id)]);
    expect((await links.listWorkbookLinks())[0]!.source).toBe("ok");
    expect(await links.pushLinks("all", relay)).toMatchObject({ pushed: 1 });

    await links.removeLink(id, relay);
    expect(workbook.names).toEqual([]);
  });
});

// Every flow that rewrites the registry queues on the one lock. The race
// suite covers auto-push against export and remove; these are the two the
// Links tab's own buttons can start on top of each other.
describe("two of the tab's own flows at once", () => {
  it("keeps a link removed while an export was still uploading", async () => {
    const first = await links.exportSelection(ws, relay);
    helpers.seed("Model!H2", [[11, 12]]);
    helpers.select("Model!H2:I2");

    const held = relay.holdNextPut();
    const exporting = links.exportSelection(ws, relay);
    await held.started;
    const removing = links.removeLink(first.id, relay);
    await settle();
    // Queued behind the export: the removal has not read the registry yet.
    expect(
      JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links,
    ).toHaveLength(1);

    held.release();
    const second = await exporting;
    await removing;

    const rows = await links.listWorkbookLinks();
    expect(rows.map((row) => row.entry.id)).toEqual([second.id]);
    expect(relay.links.has(first.id)).toBe(false);
    expect(relay.links.has(second.id)).toBe(true);
  });

  it("keeps a push's new revision when an export lands on top of it", async () => {
    const first = await links.exportSelection(ws, relay);
    helpers.seed("Model!H2", [[11, 12]]);
    helpers.select("Model!H2:I2");

    const held = relay.holdNextPut();
    const pushing = links.pushLinks("all", relay);
    await held.started;
    const exporting = links.exportSelection(ws, relay);
    await settle();

    held.release();
    expect(await pushing).toMatchObject({ pushed: 1, failed: 0 });
    const second = await exporting;

    const rows = await links.listWorkbookLinks();
    expect(rows.map((row) => row.entry.id).sort()).toEqual(
      [first.id, second.id].sort(),
    );
    // The push's revision survived the export writing the registry back.
    const pushed = rows.find((row) => row.entry.id === first.id)!.entry;
    expect(pushed.rev).toBe(2);
    expect(relay.links.get(first.id)!.rev).toBe(2);
  });
});

describe("the reads that answer nothing", () => {
  it("takes the empty string when the host cannot say what the file is", async () => {
    const office = Office as unknown as {
      context: { document: { getFilePropertiesAsync: unknown } };
    };
    const succeeded = office.context.document.getFilePropertiesAsync;
    office.context.document.getFilePropertiesAsync = (
      callback: (result: { status: string; value: { url: string } }) => void,
    ) => {
      callback({ status: "failed", value: { url: "" } });
    };
    try {
      const { id } = await links.exportSelection(ws, relay);
      expect(await payloadSource(id)).toMatchObject({ workbook: "" });
    } finally {
      office.context.document.getFilePropertiesAsync = succeeded;
    }
  });

  it("pushes an emptied cell as empty text rather than refusing", async () => {
    helpers.seed("Model!H2", [["Revenue"]]);
    helpers.select("Model!H2");
    const { id } = await links.exportSelectionAsText(ws, relay);
    helpers.seed("Model!H2", [[""]]);

    expect(await links.pushLinks("all", relay)).toMatchObject({ pushed: 1 });
    const entry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)))
      .links[0];
    const keys = await deriveLinkKeys(entry.token);
    const payload = decodePayload(
      await open(keys.enc, id, relay.links.get(id)!.blob),
    );
    expect(payload.kind === "text" && payload.text).toBe("");
  });
});
