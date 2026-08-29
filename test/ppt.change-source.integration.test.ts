// "Change source" against the fake host and the fake relay: which inbox items
// a link's row offers and in what order, the re-point itself (new picture, new
// tags, same box, inbox item consumed, the row afterwards current against the
// NEW link), and the two refusals - more than one ticked row, and an inbox with
// nothing in it. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_KEY, TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, seedLink, src } from "./ppt.support";
import type * as ChangeSourceModule from "../src/ppt/change-source";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

// Same 2:1 aspect, different bytes: the picture must be seen to change without
// the aspect rule earning the right to write a height.
const OLD = fakePng(800, 400);
const NEW = fakePng(1600, 800);

interface Tag {
  id: string;
  kind: string;
  rev: number;
  src: { workbook: string };
}

let links: typeof LinksModule;
let changeSource: typeof ChangeSourceModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  changeSource = await import("../src/ppt/change-source");
});
afterEach(() => {
  uninstallFakePpt();
});

function shape(): FakePptShape {
  return presentation.slides[0]!.shapes[0]!;
}

function tag(): Tag {
  return JSON.parse(shape().tags.get(TAG_LINK)!) as Tag;
}

// A deck holding one link from Model_v4, plus two exports waiting in the inbox:
// the same table re-exported from Model_v5, and an unrelated chart.
async function seeded(): Promise<{
  ws: Workspace;
  row: LinksModule.LinkRow;
  newer: InboxItem;
  other: InboxItem;
}> {
  const ws = await createWorkspace(memoryStore());
  const placed = await seedLink(OLD);
  await links.insertFromInbox(placed, ws, relay);
  const newer = await seedLink(NEW, "Model_v5.xlsx");
  const other: InboxItem = {
    ...(await seedLink(fakePng(10, 10), "Other.xlsx")),
    kind: "chart",
    label: "Charts: Revenue",
    src: {
      workbook: "Other.xlsx",
      sheet: "Charts",
      ref: "Revenue",
      anchor: "SMT_LINK_99999999",
    },
  };
  const row = (await links.listLinks(relay))[0]!;
  return { ws, row, newer, other };
}

describe("candidatesFor", () => {
  it("offers every waiting export, the one with the row's label first", async () => {
    const { row, newer, other } = await seeded();

    const candidates = changeSource.candidatesFor(row, [other, newer]);

    expect(candidates.map((item) => item.src.workbook)).toEqual([
      "Model_v5.xlsx",
      "Other.xlsx",
    ]);
  });

  // The label can change when a table is re-exported over a moved range, so the
  // anchor is the second signal - and an item matching neither is still offered.
  it("ranks a same-anchor export above an unrelated one", async () => {
    const { row, other } = await seeded();
    const renamed: InboxItem = {
      ...(await seedLink(NEW, "Model_v5.xlsx")),
      label: "Model!B4:F20",
      src: { ...src, workbook: "Model_v5.xlsx" },
    };

    const candidates = changeSource.candidatesFor(row, [other, renamed]);

    expect(candidates.map((item) => item.label)).toEqual([
      "Model!B4:F20",
      "Charts: Revenue",
    ]);
  });
});

describe("changeSource", () => {
  it("repaints from the new link, keeps the box, and carries the new id, token and source into the tags", async () => {
    const { ws, row, newer } = await seeded();
    const before = shape();
    const box = {
      left: before.left,
      top: before.top,
      width: before.width,
      height: before.height,
    };
    expect(tag()).toMatchObject({ id: row.found.tag.id, rev: 1 });

    const summary = await changeSource.changeSource(row, newer, ws, relay);

    expect(summary).toBe("Source changed: Model_v4.xlsx -> Model_v5.xlsx");
    expect(shape().fillImage).toBe(NEW);
    expect({
      left: shape().left,
      top: shape().top,
      width: shape().width,
      height: shape().height,
    }).toEqual(box);
    expect(tag()).toMatchObject({
      id: newer.id,
      kind: "range",
      rev: 1,
      src: { workbook: "Model_v5.xlsx" },
    });
    expect(shape().tags.get(TAG_KEY)).toBe(newer.token);
  });

  it("consumes the inbox item and reports the row current against the new link", async () => {
    const { ws, row, newer } = await seeded();
    await relay.postInbox(ws.id, ws.auth, newer.id, new Uint8Array([1]));
    expect(await relay.listInbox(ws.id, ws.auth)).toHaveLength(1);

    await changeSource.changeSource(row, newer, ws, relay);

    expect(await relay.listInbox(ws.id, ws.auth)).toHaveLength(0);
    const after = await links.listLinks(relay);
    expect(after).toHaveLength(1);
    expect(after[0]!.status).toBe("current");
    expect(after[0]!.found.tag.id).toBe(newer.id);
    // The link the picture used to track is still on the relay, untouched.
    expect(relay.links.has(row.found.tag.id)).toBe(true);
  });

  // The tags go on before the picture, so a host that refuses the repaint would
  // otherwise leave a shape claiming the new link while showing the old render
  // - and reporting itself up to date.
  it("puts the old tags back when the repaint fails, and keeps the inbox item", async () => {
    const { ws, row, newer } = await seeded();
    await relay.postInbox(ws.id, ws.auth, newer.id, new Uint8Array([1]));
    const retagged: string[] = [];
    const refusing: LinksModule.PptHost = {
      scanLinks: () => Promise.resolve([]),
      insertLink: () => Promise.reject(new Error("unused")),
      refreshLink: () =>
        Promise.reject(new Error("PowerPoint refused the picture")),
      retagLink: (_found, _tag, token) => {
        retagged.push(token);
        return Promise.resolve();
      },
      breakLink: () => Promise.resolve(),
      goToSlide: () => Promise.resolve(),
    };

    await expect(
      changeSource.changeSource(row, newer, ws, relay, refusing),
    ).rejects.toThrow(
      "change source Model!B4:F12: PowerPoint refused the picture",
    );

    expect(retagged).toEqual([newer.token, row.found.token]);
    expect(await relay.listInbox(ws.id, ws.auth)).toHaveLength(1);
  });

  it("warns when the new source is a chart and the link tracked a range", async () => {
    const { row, newer, other } = await seeded();
    expect(changeSource.kindWarning(row, newer)).toBeUndefined();
    expect(changeSource.kindWarning(row, other)).toBe(
      "The new source is a chart; this link tracked a range.",
    );
  });
});

describe("refusals", () => {
  it("refuses anything but exactly one ticked row", async () => {
    const { row } = await seeded();
    expect(() => changeSource.requireOneRow([])).toThrow("exactly one link");
    expect(() => changeSource.requireOneRow([row, row])).toThrow(
      "exactly one link",
    );
    expect(changeSource.requireOneRow([row])).toBe(row);
  });

  it("refuses an empty inbox, and an unknown choice from the list", async () => {
    const { row, newer } = await seeded();
    expect(() =>
      changeSource.requireCandidates(changeSource.candidatesFor(row, [])),
    ).toThrow("Nothing waiting in the Inbox");
    expect(() => changeSource.pickCandidate([newer], "nothing")).toThrow(
      "Choose an export",
    );
  });
});
