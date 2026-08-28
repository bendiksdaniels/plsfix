// Links inside groups, end to end against the fake host: a picture the user
// dragged into a group is still found by its tags, repainted where it sits,
// broken in place, and followed through a copy - down to three levels of
// nesting. Strict load semantics are on, so a missing load() fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import {
  encodePayload,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type Payload,
} from "../src/link/model";
import { fitToSlide } from "../src/link/status";
import { createWorkspace, type KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "SMT_LINK_00000000",
};

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

function payloadFor(png: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

async function publish(id: string, token: string, png: string): Promise<void> {
  const keys = await deriveLinkKeys(token);
  const payload = encodePayload(payloadFor(png));
  await relay.putLink(id, keys.auth, await seal(keys.enc, id, payload));
}

async function seedLink(png: string): Promise<InboxItem> {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  await publish(id, token, png);
  return {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src,
    createdAt: new Date().toISOString(),
  };
}

// A link on the first slide, then dropped into a group the way a user does:
// the picture leaves the slide and lives inside the group shape.
async function groupedLink(): Promise<{
  item: InboxItem;
  picture: FakePptShape;
  group: FakePptShape;
}> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedLink(fakePng(800, 400));
  await links.insertFromInbox(item, ws, relay);
  const slide = presentation.slides[0]!;
  const picture = slide.shapes.at(-1)!;
  const caption = presentation.addShape(slide, { left: 5, top: 5, width: 60 });
  const group = presentation.groupShapes([picture.id, caption.id], slide.id);
  return { item, picture, group };
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 2 });
  presentation = host.presentation;
  helpers = host.helpers;
  relay = new FakeRelay();
  links = await import("../src/ppt/links");
  helpers.selectSlide(presentation.slides[0]!.id);
});
afterEach(() => {
  uninstallFakePpt();
});

describe("links inside a group", () => {
  it("finds the picture with its own geometry and the path down to it", async () => {
    const { picture, group } = await groupedLink();
    const rows = await links.listLinks(relay);
    expect(rows).toHaveLength(1);
    const found = rows[0]!.found;
    expect(found.shapeId).toBe(picture.id);
    expect(found.groupPath).toEqual([group.id]);
    expect(found.slideIndex).toBe(0);
    expect({
      left: found.left,
      top: found.top,
      width: found.width,
      height: found.height,
    }).toEqual(fitToSlide(800, 400));
    expect(rows[0]!.status).toBe("current");
  });

  it("repaints it in place and leaves it in the group", async () => {
    const { item, picture, group } = await groupedLink();
    await publish(item.id, item.token, fakePng(800, 400));
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(picture.setImageCalls).toBe(2);
    expect(group.group!.shapes).toContain(picture);
    expect(presentation.slides[0]!.shapes).toEqual([group]);
    expect(JSON.parse(picture.tags.get(TAG_LINK)!).rev).toBe(2);
  });

  it("breaks the link where it sits, keeping the picture in the group", async () => {
    const { picture, group } = await groupedLink();
    const [row] = await links.listLinks(relay);
    const host = await import("../src/ppt/host");
    await host.breakLink(row!.found);
    expect(picture.tags.size).toBe(0);
    expect(picture.fillImage).toBe(fakePng(800, 400));
    expect(group.group!.shapes).toContain(picture);
    expect(await links.listLinks(relay)).toHaveLength(0);
  });

  // Copying a group copies every shape in it, new ids and all: the two links
  // are then told apart only by the path each was found through.
  it("follows a copy of the group to its own picture", async () => {
    const { item, picture, group } = await groupedLink();
    const copy = presentation.copyShape(group.id, presentation.slides[1]!.id);
    const twin = copy.group!.shapes[0]!;
    expect(twin.id).not.toBe(picture.id);
    await publish(item.id, item.token, fakePng(800, 400));
    const rows = await links.listLinks(relay);
    expect(rows.map((row) => row.found.groupPath)).toEqual([
      [group.id],
      [copy.id],
    ]);
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect([picture.setImageCalls, twin.setImageCalls]).toEqual([2, 2]);
  });

  it("walks a group inside a group", async () => {
    const { item, picture, group } = await groupedLink();
    const slide = presentation.slides[0]!;
    const outer = presentation.groupShapes([group.id], slide.id);
    await publish(item.id, item.token, fakePng(800, 400));
    const rows = await links.listLinks(relay);
    expect(rows[0]!.found.groupPath).toEqual([outer.id, group.id]);
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(picture.setImageCalls).toBe(2);
  });

  // The cap: three levels of groups are walked, a fourth is somebody's art.
  it("finds a picture three groups deep and ignores one four deep", async () => {
    const near = await groupedLink();
    const far = await groupedLink();
    const slide = presentation.slides[0]!;
    let nearGroup = near.group;
    let farGroup = far.group;
    for (let level = 0; level < 2; level += 1) {
      nearGroup = presentation.groupShapes([nearGroup.id], slide.id);
      farGroup = presentation.groupShapes([farGroup.id], slide.id);
    }
    farGroup = presentation.groupShapes([farGroup.id], slide.id);
    const rows = await links.listLinks(relay);
    expect(rows.map((row) => row.found.shapeId)).toEqual([near.picture.id]);
    expect(rows[0]!.found.groupPath).toHaveLength(3);
  });

  // A grouped picture is never reinserted: that would drop it onto the slide,
  // out of the group the user put it in.
  it("fails a grouped row on a host below PowerPointApi 1.8", async () => {
    const { item, picture } = await groupedLink();
    await publish(item.id, item.token, fakePng(800, 400));
    const rows = await links.listLinks(relay);
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([
      "refresh Model!B4:F12: grouped pictures need PowerPoint 2504/16.96 or newer",
    ]);
    expect(picture.setImageCalls).toBe(1);
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(picture.tags.get(TAG_LINK)!).rev).toBe(1);
  });

  it("skips groups entirely on a host below PowerPointApi 1.8", async () => {
    await groupedLink();
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
    expect(await links.listLinks(relay)).toHaveLength(0);
  });
});

describe("the fake host's groups", () => {
  it("polices group loads and keeps a slide's shapes apart from a group's", async () => {
    const { picture, group } = await groupedLink();
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const shapes = slides.items[0]!.shapes;
      const inside = shapes.getItem(group.id).group.shapes;
      expect(() => inside.items).toThrow(/PropertyNotLoaded/);
      inside.load("items/id");
      await context.sync();
      expect(inside.items.map((shape) => shape.id)).toContain(picture.id);
      expect(() => shapes.getItem(picture.id)).toThrow(/ItemNotFound/);
    });
  });
});
