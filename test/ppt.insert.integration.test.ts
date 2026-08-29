// Inbox insert, end to end against the fake host and the fake relay:
// pulling an item out of the inbox creates a tagged, picture-filled
// rectangle sized to the image and clears the inbox; a blob sealed under
// another workspace's key is skipped. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { seal } from "../src/link/crypto";
import { encodeInboxItem, TAG_KEY, TAG_LINK } from "../src/link/model";
import { fitToSlide } from "../src/link/status";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
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

describe("insert from inbox", () => {
  it("creates a tagged picture-filled rectangle sized to the image and clears the inbox", async () => {
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
    expect({
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    }).toEqual(fitToSlide(800, 400));
    expect(JSON.parse(shape.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "range",
      rev: 1,
    });
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(await links.listInbox(ws, relay)).toHaveLength(0);
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
