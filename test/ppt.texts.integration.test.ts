// Text links against the fake host and the fake relay: an insert lands one
// tagged, auto-sized text box in the slide's free space, an update writes the
// text and the tag and nothing else, a revert puts the older text back, and a
// break leaves a plain text box behind. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { overlaps } from "../src/layout";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
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
  pushText,
  seedLink,
  seedText,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type * as RevertModule from "../src/ppt/revert";

enableStrictLoadSemantics();

const TEXT = "EUR 15.7m";
const AFTER = "EUR 16.1m";

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

async function insertText(
  text = TEXT,
): Promise<Awaited<ReturnType<typeof seedText>>> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedText(text);
  await links.insertFromInbox(item, ws, relay);
  return item;
}

function shape(index = 0): FakePptShape {
  return presentation.slides[0]!.shapes[index]!;
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

describe("insert a text link", () => {
  it("lands a tagged, auto-sized text box in the free space", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedText(TEXT);
    const placed = await links.insertFromInbox(item, ws, relay);

    const inserted = shape();
    expect(inserted.id).toBe(placed.shapeId);
    expect(inserted.type).toBe("TextBox");
    expect(inserted.name).toBe(`pls,fix text ${item.label}`);
    expect(inserted.text).toBe(TEXT);
    expect(inserted.autoSize).toBe("AutoSizeShapeToFitText");
    expect(inserted.wordWrap).toBe(false);
    expect(JSON.parse(inserted.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "text",
      rev: 1,
    });
    expect(inserted.tags.get(TAG_KEY)).toBe(item.token);
    expect(placed.overlapping).toBe(false);
    expect(
      await links.listInbox(await createWorkspace(memoryStore()), relay),
    ).toHaveLength(0);
  });

  it("lands a second text box beside the first, not on it", async () => {
    await insertText();
    await insertText("EUR 3.2m");
    expect(overlaps(box(shape(0)), box(shape(1)))).toBe(false);
  });

  it("costs no more syncs than a picture insert", async () => {
    const ws = await createWorkspace(memoryStore());
    const picture = await seedLink(fakePng(800, 400));
    const before = helpers.syncCount();
    await links.insertFromInbox(picture, ws, relay);
    const pictureSyncs = helpers.syncCount() - before;

    const item = await seedText(TEXT);
    const start = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    expect(helpers.syncCount() - start).toBeLessThanOrEqual(pictureSyncs);
  });
});

describe("update a text link", () => {
  it("writes the text in place and leaves the box where it was", async () => {
    const item = await insertText();
    const moved = shape();
    moved.left = 100;
    moved.top = 200;
    moved.width = 150;
    moved.height = 40;
    await pushText(item, AFTER);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(shape().id).toBe(moved.id);
    expect(shape().text).toBe(AFTER);
    expect(box(shape())).toEqual({
      left: 100,
      top: 200,
      width: 150,
      height: 40,
    });
    expect(JSON.parse(shape().tags.get(TAG_LINK)!).rev).toBe(2);
  });

  it("puts the older text back on a revert", async () => {
    const revert: typeof RevertModule = await import("../src/ppt/revert");
    const item = await insertText();
    await pushText(item, AFTER);
    await links.updateLinks(await links.listLinks(relay), relay);

    const summary = await revert.revertLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ reverted: 1, failed: 0 });
    expect(shape().text).toBe(TEXT);
    expect(JSON.parse(shape().tags.get(TAG_LINK)!).rev).toBe(1);
  });
});

describe("break a text link", () => {
  it("drops the tags and leaves a plain text box", async () => {
    await insertText();
    const [row] = await links.listLinks(relay);
    const host = await import("../src/ppt/host");
    await host.breakLink(row!.found);
    expect(shape().text).toBe(TEXT);
    expect(shape().tags.size).toBe(0);
    expect(await links.listLinks(relay)).toHaveLength(0);
  });
});
