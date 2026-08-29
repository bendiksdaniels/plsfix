// Inbox insert, end to end against the fake host and the fake relay:
// pulling an item out of the inbox creates a tagged, picture-filled
// rectangle sized to the image, placed in the slide's free space, and clears
// the inbox; a blob sealed under another workspace's key is skipped. Strict
// load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { seal } from "../src/link/crypto";
import { overlaps, type Box } from "../src/layout";
import { encodeInboxItem, TAG_KEY, TAG_LINK } from "../src/link/model";
import { fitToSlide } from "../src/link/status";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, seedLink } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

function box(shape: FakePptShape): Box {
  return {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

describe("insert from inbox", () => {
  it("creates a tagged picture-filled rectangle sized to the image, at the first free spot, and clears the inbox", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await relay.postInbox(
      ws.id,
      ws.auth,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );
    expect(await links.listInbox(ws, relay)).toHaveLength(1);
    await links.insertFromInbox(item, ws, relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect(shape.fillImage).toBe(fakePng(800, 400));
    expect(shape.lineVisible).toBe(false);
    const fitted = fitToSlide(800, 400);
    // The size fitToSlide chose, at the first free spot - the margin corner of
    // an empty slide, not the centre the picture used to be dropped on.
    expect({
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    }).toEqual({
      left: 36,
      top: 36,
      width: fitted.width,
      height: fitted.height,
    });
    expect(JSON.parse(shape.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "range",
      rev: 1,
    });
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(await links.listInbox(ws, relay)).toHaveLength(0);
  });

  // Two exports onto one slide used to land on exactly the same centred box,
  // so the second hid the first. Placement reads what the slide already holds.
  it("places a second insert clear of the first", async () => {
    const ws = await createWorkspace(memoryStore());
    for (const size of [800, 400]) {
      const item = await seedLink(fakePng(size, size / 2));
      await relay.postInbox(
        ws.id,
        ws.auth,
        item.id,
        await seal(ws.enc, ws.id, encodeInboxItem(item)),
      );
      await links.insertFromInbox(item, ws, relay);
    }
    const [first, second] = presentation.slides[0]!.shapes;
    expect(second).toBeDefined();
    expect(overlaps(box(first!), box(second!))).toBe(false);
  });

  // A slide with no room left still takes the picture - centred, over what is
  // there - and says so, because refusing the insert helps nobody.
  it("reports the overlap when the slide has no free space", async () => {
    const ws = await createWorkspace(memoryStore());
    presentation.addShape(presentation.slides[0]!, {
      left: 0,
      top: 0,
      width: 960,
      height: 540,
    });
    const item = await seedLink(fakePng(800, 400));
    await relay.postInbox(
      ws.id,
      ws.auth,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );
    expect(await links.insertFromInbox(item, ws, relay)).toMatchObject({
      overlapping: true,
    });
  });

  // An empty layout placeholder is the slide's own furniture: a picture is
  // meant to land on it, and one that holds text is not to be covered.
  it("places over an empty placeholder but not over one with text", async () => {
    const ws = await createWorkspace(memoryStore());
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      type: "Placeholder",
      left: 36,
      top: 36,
      width: 888,
      height: 100,
    });
    const item = await seedLink(fakePng(800, 400));
    await relay.postInbox(
      ws.id,
      ws.auth,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );
    await links.insertFromInbox(item, ws, relay);
    expect(slide.shapes[1]!.top).toBe(36);

    slide.shapes[0]!.hasText = true;
    const second = await seedLink(fakePng(800, 400));
    await relay.postInbox(
      ws.id,
      ws.auth,
      second.id,
      await seal(ws.enc, ws.id, encodeInboxItem(second)),
    );
    await links.insertFromInbox(second, ws, relay);
    expect(slide.shapes[2]!.top).toBeGreaterThan(36);
  });

  // The reason server/src/store.rs keys inbox_v2 on (ws, id, auth_hash) and no
  // longer on (ws, id): a foreign key that announces the same link id gets its
  // own row rather than the pane's slot. On the old key it destroyed the item
  // the deck was waiting for, and the pane could not tell it had ever existed.
  it("keeps a foreign key's announcement of the same link out of ours", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(10, 10));
    await relay.postInbox(
      ws.id,
      ws.auth,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );
    // Same workspace, same link id, another writer's key - a re-keyed workbook
    // exporting the link it already exported once.
    const foreign = "F".repeat(43);
    await relay.postInbox(
      ws.id,
      foreign,
      item.id,
      await seal(ws.enc, ws.id, encodeInboxItem(item)),
    );

    expect(relay.inbox.size).toBe(2);
    expect((await links.listInbox(ws, relay)).map((row) => row.id)).toEqual([
      item.id,
    ]);
    expect(await relay.listInbox(ws.id, foreign)).toHaveLength(1);
    // And ours is still the one the deck consumes and clears.
    await links.insertFromInbox(item, ws, relay);
    expect(await links.listInbox(ws, relay)).toHaveLength(0);
    expect(await relay.listInbox(ws.id, foreign)).toHaveLength(1);
  });

  it("skips an inbox blob sealed with another workspace key, and an empty deck summarizes to nothing", async () => {
    const ws = await createWorkspace(memoryStore());
    const other = await createWorkspace(memoryStore());
    const mine = await seedLink(fakePng(10, 10));
    const theirs = await seedLink(fakePng(10, 10));
    await relay.postInbox(
      ws.id,
      ws.auth,
      mine.id,
      await seal(ws.enc, ws.id, encodeInboxItem(mine)),
    );
    await relay.postInbox(
      ws.id,
      ws.auth,
      theirs.id,
      await seal(other.enc, other.id, encodeInboxItem(theirs)),
    );
    const inbox = await links.listInbox(ws, relay);
    expect(inbox.map((entry) => entry.id)).toEqual([mine.id]);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(links.summarize(summary)).toBe("No links found");
  });
});
