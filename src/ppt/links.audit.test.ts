// Slice F audit, the link flows below the pane: a group the user pulled
// apart under a link, the per-row GET the deferred path falls back to, a
// change of source whose rollback fails too, a host that rejects with
// something that is not an Error, and the column widths a table with no
// widths at all is built with. Over the shared fake deck and fake relay.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_LINK } from "../link/model";
import { RelayError } from "../link/relay";
import { createWorkspace } from "../link/workspace";
import type { FakeRelay } from "../../test/fakerelay";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "../../test/fakeppt";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  pushText,
  seedLink,
  seedTable,
  seedText,
} from "../../test/ppt.support";
import { CONTENT_WIDTH } from "./placement";
import { columnWidths, tableSize } from "./tables";
import type * as ChangeSourceModule from "./change-source";
import type * as LinksModule from "./links";

enableStrictLoadSemantics();

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

describe("a group the user pulled apart", () => {
  // The rows a button acts on are the last scan's, and their path runs down
  // through the groups the link sat in. Ungrouping moves the picture back onto
  // the slide, so that path no longer reaches it: the row is named as a
  // failure, nothing else in the deck is touched, and the rescan every action
  // ends with finds it again - the next click updates it.
  it("names the row, and the rescan behind it makes the next click work", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const other = await seedLink(fakePng(400, 200));
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(other, ws, relay);
    const picture = shape();
    const caption = presentation.addShape(presentation.slides[0]!, {
      left: 5,
      top: 5,
      width: 60,
    });
    const group = presentation.groupShapes(
      [picture.id, caption.id],
      presentation.slides[0]!.id,
    );
    // The scan lists the slides' own shapes first and the groups' children
    // after them, so the grouped link is the one carrying a path.
    const stale = await links.listLinks(relay);
    const grouped = stale.find((row) => row.found.groupPath !== undefined)!;
    expect(grouped.found.groupPath).toEqual([group.id]);
    // Ctrl+Shift+G: the members go back onto the slide and the group is gone.
    const slide = presentation.slides[0]!;
    slide.shapes.splice(slide.shapes.indexOf(group), 1, ...group.group!.shapes);
    for (const one of [item, other]) await pushAgain(one, fakePng(800, 400));

    const summary = await links.updateLinks(stale, relay);

    expect(summary).toMatchObject({ updated: 1, failed: 1 });
    expect(summary.failures[0]).toContain("Model!B4:F12");
    expect(picture.setImageCalls).toBe(1);

    // The fresh scan reaches it on the slide, where the user left it.
    const fresh = await links.listLinks(relay);
    expect(fresh.every((row) => row.found.groupPath === undefined)).toBe(true);
    const again = await links.updateLinks(fresh, relay);
    expect(again).toMatchObject({ updated: 1, current: 1, failed: 0 });
    expect(picture.setImageCalls).toBe(2);
  });
});

describe("the per-row GET behind the batch", () => {
  it("counts a row the relay calls unchanged without repainting it", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    expect(rows[0]!.status).toBe("current");
    // An older relay knows no batch route at all, so every row falls back to
    // the GET - which answers with the revision the deck already holds.
    vi.spyOn(relay, "fetchLinks").mockRejectedValue(new Error("no such route"));
    const getLink = vi.spyOn(relay, "getLink");

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(getLink).toHaveBeenCalledTimes(1);
    expect(shape().setImageCalls).toBe(1);
  });

  it("paints nothing at all for an empty batch", async () => {
    const host = await import("./host");
    const before = helpers.syncCount();
    expect(await host.refreshLinks([])).toBe(true);
    expect(helpers.syncCount()).toBe(before);
  });
});

describe("a change of source that cannot be undone", () => {
  // The rollback is best effort: the repaint's failure is the one the user
  // needs to read, so a retag that fails too must not replace it.
  it("keeps the repaint's reason when the rollback fails as well", async () => {
    const changeSource: typeof ChangeSourceModule =
      await import("./change-source");
    const ws = await createWorkspace(memoryStore());
    const placed = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;
    const newer = await seedLink(fakePng(1600, 800), "Model_v5.xlsx");
    let retags = 0;
    const failing: LinksModule.PptHost = {
      scanLinks: () => Promise.resolve([]),
      insertLink: () => Promise.reject(new Error("unused")),
      refreshLink: () => Promise.reject(new Error("PowerPoint said no")),
      retagLink: () => {
        retags += 1;
        return retags === 1
          ? Promise.resolve()
          : Promise.reject(new Error("and the rollback said no"));
      },
      breakLink: () => Promise.resolve(),
      goToSlide: () => Promise.resolve(),
    };

    await expect(
      changeSource.changeSource(row, newer, ws, relay, failing),
    ).rejects.toThrow("change source Model!B4:F12: PowerPoint said no");
    expect(retags).toBe(2);
    // The export is still waiting, so the user can try again.
    expect(await relay.listInbox(ws.id, ws.auth)).toHaveLength(0);
  });
});

describe("a host that rejects with something that is not an Error", () => {
  it("still names the link and the stage on a text refresh", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedText("EUR 15.7m");
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 16.1m");
    const rows = await links.listLinks(relay);
    // Office.js can reject a batch with a plain object rather than an Error.
    helpers.failNextSync({ code: "GeneralException" } as unknown as Error);

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    // The stage and the link survive whatever the host rejected with.
    expect(summary.failures[0]).toContain("refresh Model!B4:F12 text: ");
  });
});

describe("a table whose source reports no column widths", () => {
  // Excel reports zero for a column that is hidden, so a range of hidden
  // columns arrives with nothing to scale by. columnWidths already divides
  // evenly rather than by zero; the table's own width has to survive it too,
  // or the slide gets a table zero points wide that cannot be seen or picked
  // up again.
  const payload = {
    v: 1 as const,
    kind: "table" as const,
    rows: 1,
    cols: 4,
    cells: [[{ t: "a" }, { t: "b" }, { t: "c" }, { t: "d" }]],
    widths: [0, 0, 0, 0],
    src: {
      workbook: "Model.xlsx",
      sheet: "Model",
      ref: "A1:D1",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };

  it("falls back to the slide's content width, in even columns", () => {
    expect(tableSize(payload).width).toBe(CONTENT_WIDTH);
    expect(columnWidths(payload, tableSize(payload).width)).toEqual(
      new Array<number>(4).fill(CONTENT_WIDTH / 4),
    );
  });

  it("lands a table the user can actually see", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable([[{ t: "a" }, { t: "b" }]], [0, 0]);

    await links.insertFromInbox(item, ws, relay);

    expect(shape().type).toBe("Table");
    expect(shape().width).toBeGreaterThan(0);
  });
});

describe("update all with one row failing", () => {
  it("bumps the revisions of the rows that landed and names the one that did not", async () => {
    const ws = await createWorkspace(memoryStore());
    const items = [];
    for (const slide of presentation.slides) {
      helpers.selectSlide(slide.id);
      const item = await seedLink(fakePng(10, 10));
      await links.insertFromInbox(item, ws, relay);
      items.push(item);
    }
    for (const item of items) await pushAgain(item, fakePng(20, 20));
    const rows = await links.listLinks(relay);
    // The middle shape is gone from the deck, so its repaint - batched or
    // not - throws while its neighbours repaint.
    presentation.deleteShape(rows[1]!.found.shapeId);

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 2, failed: 1 });
    expect(summary.failures).toHaveLength(1);
    const revs = presentation.slides
      .flatMap((slide) => slide.shapes)
      .map(
        (one) => (JSON.parse(one.tags.get(TAG_LINK)!) as { rev: number }).rev,
      );
    expect(revs).toEqual([2, 2]);
  });
});

// Found by slice E2: "tooLarge" was the one RelayErrorKind with no consumer,
// so a 413 - the relay's own 4 MiB link route, or an nginx or Cloudflare hop
// in front of it - reached the user as a status line nobody can act on.
describe("a refusal because the export is too big", () => {
  const TOO_BIG =
    "That export is too big to send. Export a smaller range from Excel.";
  const refused = (): RelayError =>
    new RelayError(
      "tooLarge",
      "relay GET /api/links/aaaa: 413 payload too large",
      413,
    );

  it("says so in words in the update summary", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(1600, 800));
    const rows = await links.listLinks(relay);
    vi.spyOn(relay, "fetchLinks").mockRejectedValue(refused());
    vi.spyOn(relay, "getLink").mockRejectedValue(refused());

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([`Model!B4:F12: ${TOO_BIG}`]);
  });

  it("says so in words on an insert", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    vi.spyOn(relay, "getLink").mockRejectedValue(refused());

    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
      `insert Model!B4:F12: ${TOO_BIG}`,
    );
  });

  it("says so in words on a revert", async () => {
    const revert = await import("./revert");
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(1600, 800));
    await links.updateLinks(await links.listLinks(relay), relay);
    const rows = await links.listLinks(relay);
    vi.spyOn(relay, "getLinkRev").mockRejectedValue(refused());

    const summary = await revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 0, failed: 1 });
    expect(summary.failures).toEqual([`Model!B4:F12: ${TOO_BIG}`]);
  });
});
