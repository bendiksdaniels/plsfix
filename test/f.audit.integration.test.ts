// Slice F audit: the insert / update / revert / change-source paths the
// existing ppt.* suites leave uncovered - a table the user edited under the
// link, a grouped table whose grid moved, a text refresh the host refuses, an
// insert with no slide selected, a revert whose repaint fails, the batch
// fetch's fallbacks, and the candidate rule "Change source" has to keep.
// Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { newToken } from "../src/link/crypto";
import { TAG_KEY, TAG_LINK, type TableCell } from "../src/link/model";
import { RelayError } from "../src/link/relay";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  pushTable,
  pushText,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";
import type * as ChangeSourceModule from "../src/ppt/change-source";
import type * as LinksModule from "../src/ppt/links";
import type * as RevertModule from "../src/ppt/revert";

enableStrictLoadSemantics();

const WIDTHS = [80, 60];
const CELLS: TableCell[][] = [
  [{ t: "Revenue" }, { t: "1 000" }],
  [{ t: "Costs" }, { t: "-400" }],
];

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

function shape(index = 0): FakePptShape {
  return presentation.slides[0]!.shapes[index]!;
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

// A host that refuses every repaint, for the paths whose failure line is the
// thing under test.
const refusing: LinksModule.PptHost = {
  scanLinks: () => Promise.resolve([]),
  insertLink: () => Promise.reject(new Error("unused")),
  refreshLink: () => Promise.reject(new Error("PowerPoint refused the shape")),
  retagLink: () => Promise.resolve(),
  breakLink: () => Promise.resolve(),
  goToSlide: () => Promise.resolve(),
};

describe("a table the user edited under the link", () => {
  it("rebuilds it at the corner the user left it at when a row was deleted", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const before = shape();
    before.left = 120;
    before.top = 60;
    before.width = 260;
    // The user deleted the second row in PowerPoint; the source still has two.
    before.table!.cells.pop();
    before.table!.rowCount = 1;
    await pushTable(item, CELLS, WIDTHS);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    const after = shape();
    expect(after.id).not.toBe(before.id);
    expect([after.left, after.top, after.width]).toEqual([120, 60, 260]);
    expect(after.table!.rowCount).toBe(2);
    expect(
      after.table!.cells.map((row) => row.map((cell) => cell.text)),
    ).toEqual([
      ["Revenue", "1 000"],
      ["Costs", "-400"],
    ]);
    expect(after.tags.get(TAG_KEY)).toBe(item.token);
  });

  // The rebuild drops the new table on the slide, not back into the group, so
  // a grouped table is left alone and the row says what to do about it.
  it("refuses to rebuild a table the user grouped", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const table = shape();
    const caption = presentation.addShape(presentation.slides[0]!, {
      left: 5,
      top: 5,
      width: 60,
    });
    presentation.groupShapes(
      [table.id, caption.id],
      presentation.slides[0]!.id,
    );
    await pushTable(item, [...CELLS, [{ t: "Profit" }, { t: "600" }]], WIDTHS);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([
      "refresh Model!B4:F12 table: ungroup the table before its size can change",
    ]);
    expect(table.table!.rowCount).toBe(2);
  });

  // Geometry stays the user's while the grid still matches: the repaint is
  // cells only, whatever the source's own column widths now say.
  it("leaves a moved and resized table exactly where it is", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const moved = shape();
    moved.left = 33;
    moved.top = 44;
    moved.width = 555;
    moved.height = 66;
    await pushTable(item, CELLS, [400, 20]);

    await links.updateLinks(await links.listLinks(relay), relay);

    expect(box(shape())).toEqual({
      left: 33,
      top: 44,
      width: 555,
      height: 66,
    });
  });
});

describe("a text link the host refuses", () => {
  it("names the link and the reason in one line", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedText("EUR 15.7m");
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    await pushText(item, "EUR 16.1m");
    // The user deleted the text box between the scan and the repaint.
    presentation.deleteShape(rows[0]!.found.shapeId);

    const summary = await links.updateLinks(
      await links.listLinks(relay).then(() => rows),
      relay,
    );

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain("refresh Model!B4:F12 text: ");
    // The reason used to be office.js's own "ItemNotFound: no shape ..."; the
    // P3 stress pass turned the one host code a stale row can earn into the
    // pane's sentence (src/ppt/missing-shape.ts).
    expect(summary.failures[0]).toContain(
      "that object is no longer where the list had it",
    );
  });
});

describe("an insert with no slide selected", () => {
  it("says so for a picture, a table and a text, and writes nothing", async () => {
    const ws = await createWorkspace(memoryStore());
    const waiting = [
      await seedLink(fakePng(800, 400)),
      await seedTable(CELLS, WIDTHS),
      await seedText("EUR 15.7m"),
    ];
    helpers.clearSelection();

    for (const item of waiting) {
      await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
        `insert ${item.label}: select a slide first.`,
      );
    }
    expect(presentation.slides.flatMap((slide) => slide.shapes)).toEqual([]);
  });

  it("reports a relay that answered no picture at all", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    vi.spyOn(relay, "getLink").mockResolvedValue("unchanged");

    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
      `insert ${item.label}: the relay returned no picture.`,
    );
  });
});

describe("a revert the host refuses", () => {
  it("counts the row and names it", async () => {
    const revert: typeof RevertModule = await import("../src/ppt/revert");
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(1600, 800));
    await links.updateLinks(await links.listLinks(relay), relay);
    const rows = await links.listLinks(relay);

    const summary = await revert.revertLinks(rows, relay, refusing);

    expect(summary).toMatchObject({ reverted: 0, noPrevious: 0, failed: 1 });
    expect(summary.failures).toEqual([
      "Model!B4:F12: PowerPoint refused the shape",
    ]);
    // Nothing was repainted, so the deck still holds the newer revision.
    expect(shape().fillImage).toBe(fakePng(1600, 800));
    expect(JSON.parse(shape().tags.get(TAG_LINK)!).rev).toBe(2);
  });
});

describe("the batch fetch's fallbacks", () => {
  it("asks for a copy whose revision drifted, and counts an unopenable blob", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const original = presentation.slides[0]!.shapes[0]!;
    const copy = presentation.copyShape(
      original.id,
      presentation.slides[1]!.id,
    );
    // One copy still claims the revision before the other's: the batch may not
    // send a knownRev at all, so the relay answers with the picture.
    const tag: { rev: number } = JSON.parse(copy.tags.get(TAG_LINK)!);
    copy.tags.set(TAG_LINK, JSON.stringify({ ...tag, rev: tag.rev - 1 }));
    await pushAgain(item, fakePng(1600, 800));

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect([original.fillImage, copy.fillImage]).toEqual([
      fakePng(1600, 800),
      fakePng(1600, 800),
    ]);
  });

  it("fails only the rows whose blob will not open", async () => {
    const ws = await createWorkspace(memoryStore());
    const good = await seedLink(fakePng(10, 10));
    const bad = await seedLink(fakePng(20, 20));
    for (const item of [good, bad]) {
      await links.insertFromInbox(item, ws, relay);
    }
    for (const item of [good, bad]) await pushAgain(item, fakePng(30, 30));
    // Corrupt what the relay holds for one link only.
    relay.links.get(bad.id)!.blob = new Uint8Array([1, 2, 3, 4, 5]);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 1 });
    expect(summary.failures[0]).toContain("Model!B4:F12: ");
  });

  it("ignores an answer for a link it never asked about", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(20, 20));
    const rows = await links.listLinks(relay);
    const real = relay.fetchLinks.bind(relay);
    vi.spyOn(relay, "fetchLinks").mockImplementation(async (queries) => {
      const answer = await real(queries);
      answer.items.push({
        id: "not asked about",
        rev: 9,
        blob: new Uint8Array(),
      });
      answer.omitted.push({ id: "not asked about either", reason: "missing" });
      return answer;
    });

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 1, missing: 0, failed: 0 });
  });

  it("counts a link the relay lost and one whose key was changed", async () => {
    const ws = await createWorkspace(memoryStore());
    const gone = await seedLink(fakePng(10, 10));
    const rekeyed = await seedLink(fakePng(10, 10));
    for (const item of [gone, rekeyed]) {
      await links.insertFromInbox(item, ws, relay);
    }
    const rows = await links.listLinks(relay);
    vi.spyOn(relay, "fetchLinks").mockRejectedValue(
      new RelayError("server", "relay POST /api/links/fetch: 500", 500),
    );
    vi.spyOn(relay, "getLink").mockImplementation(async (id) => {
      if (id === gone.id) throw new RelayError("missing", "not found", 404);
      throw new RelayError("auth", "forbidden", 403);
    });

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({
      updated: 0,
      missing: 1,
      wrongKey: 1,
      failed: 0,
    });
    expect(links.summarize(summary)).toBe("1 missing, 1 wrong key");
  });

  it("counts a row whose key will not derive without calling the relay", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await links.insertFromInbox(item, ws, relay);
    presentation.slides[0]!.shapes[0]!.tags.set(TAG_KEY, "not a token");
    const rows = await links.listLinks(relay);
    const fetchLinks = vi.spyOn(relay, "fetchLinks");
    const getLink = vi.spyOn(relay, "getLink");

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ wrongKey: 1, updated: 0, failed: 0 });
    expect(fetchLinks).not.toHaveBeenCalled();
    expect(getLink).not.toHaveBeenCalled();
  });
});

describe("Change source candidates", () => {
  it("keeps a picture export away from a table link, and the reverse", async () => {
    const changeSource: typeof ChangeSourceModule =
      await import("../src/ppt/change-source");
    const ws = await createWorkspace(memoryStore());
    const table = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(table, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    const range = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(range, ws, relay);
    const inbox = [
      await seedLink(fakePng(1600, 800), "Model_v9.xlsx"),
      await seedTable(CELLS, [90, 70]),
      await seedText("EUR 15.7m"),
    ];
    const rows = await links.listLinks(relay);

    // A native table cannot take a picture fill in place - it would keep its
    // cells and grow an image behind them - and a picture-filled rectangle is
    // no table: getTable throws on it. Neither crossing is offered.
    expect(
      changeSource.candidatesFor(rows[0]!, inbox).map((one) => one.kind),
    ).toEqual(["table"]);
    expect(
      changeSource.candidatesFor(rows[1]!, inbox).map((one) => one.kind),
    ).toEqual(["range"]);
  });

  it("keeps the inbox order between equally good candidates", async () => {
    const changeSource: typeof ChangeSourceModule =
      await import("../src/ppt/change-source");
    const ws = await createWorkspace(memoryStore());
    const placed = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;
    // Two exports from other workbooks, neither matching the label or the
    // anchor: the inbox's own order (newest first) is the tie-break.
    const first = await seedLink(fakePng(10, 10), "A.xlsx");
    const second = await seedLink(fakePng(20, 20), "B.xlsx");
    for (const item of [first, second]) {
      item.label = "Something else";
      item.src = { ...item.src, anchor: newToken().slice(0, 20) };
    }

    expect(
      changeSource
        .candidatesFor(row, [second, first])
        .map((one) => one.src.workbook),
    ).toEqual(["B.xlsx", "A.xlsx"]);
  });
});
