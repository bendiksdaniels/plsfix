// pasteLinks against a real fakeppt deck and a fresh, memory-backed
// LocalStore: bundles are built the way Excel's LocalCollector records them
// (src/link/local-collector.ts), sealed with the real crypto, so every rule
// here is proven against the same bytes Excel would actually copy.

import { afterEach, describe, expect, it } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { localWorkspace } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import {
  encodeInboxItem,
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
} from "../src/link/model";
import { insertFromInbox, listInbox } from "../src/ppt/links";
import { pasteLinks, summarizePaste } from "../src/ppt/paste-links";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";

enableStrictLoadSemantics();

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

function picture(png: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

// One Excel export the way LocalCollector records it: the link's sealed
// payload, plus its inbox row sealed with the fixed local workspace (design
// 2026-09-23 s.4 - "an export copies one link plus its inbox row").
async function exportLink(
  collector: LocalCollector,
  id: string,
  token: string,
  payload: Payload,
  currentRev = 0,
): Promise<InboxItem> {
  const keys = await deriveLinkKeys(token);
  await collector.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
    currentRev,
  );
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: SRC,
    createdAt: new Date().toISOString(),
  };
  const localWs = await localWorkspace();
  await collector.postInbox(
    localWs.id,
    localWs.auth,
    id,
    await seal(localWs.enc, localWs.id, encodeInboxItem(item)),
  );
  return item;
}

// A shape already on slide 2, tagged at `rev`: the deck a bundle is about to
// update, the way a colleague's deck arrives.
function plantShape(
  presentation: FakePresentation,
  id: string,
  token: string,
  rev: number,
): void {
  const shape = presentation.addShape(presentation.slides[1]!, {
    type: "GeometricShape",
    fillImage: fakePng(800, 400),
    left: 100,
    top: 80,
    width: 400,
    height: 200,
  });
  const tag: LinkTag = {
    v: 1,
    id,
    kind: "range",
    rev,
    src: SRC,
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
}

let presentation: FakePresentation;

function boot(): void {
  uninstallFakePpt();
  const host = installFakePpt({ slides: 3 });
  presentation = host.presentation;
  host.helpers.selectSlide(presentation.slides[0]!.id);
}

afterEach(() => {
  uninstallFakePpt();
});

function freshStore(): Promise<LocalStore> {
  return LocalStore.open(memoryPersistence());
}

describe("pasteLinks", () => {
  it("repaints a link the deck already holds at an older revision", async () => {
    boot();
    const id = newId();
    const token = newToken();
    plantShape(presentation, id, token, 1);
    const collector = new LocalCollector();
    await exportLink(collector, id, token, picture(fakePng(400, 200)), 1);
    const store = await freshStore();

    const summary = await pasteLinks(collector.bundle(), store);

    expect(summary).toMatchObject({
      links: 1,
      updated: 1,
      current: 0,
      waiting: 0,
      failed: 0,
    });
    expect(summarizePaste(summary)).toBe("Pasted 1 link: 1 updated.");
    expect(presentation.slides[1]!.shapes[0]!.fillImage).toBe(
      fakePng(400, 200),
    );
  });

  it("says up to date the second time the same bundle is pasted", async () => {
    boot();
    const id = newId();
    const token = newToken();
    plantShape(presentation, id, token, 1);
    const collector = new LocalCollector();
    await exportLink(collector, id, token, picture(fakePng(400, 200)), 1);
    const store = await freshStore();
    const bundle = collector.bundle();
    await pasteLinks(bundle, store);

    const summary = await pasteLinks(bundle, store);

    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(summarizePaste(summary)).toBe("Pasted 1 link: 1 up to date.");
  });

  it("puts a new link in the Inbox untouched on the deck, and it can be inserted from there", async () => {
    boot();
    const id = newId();
    const token = newToken();
    const collector = new LocalCollector();
    const item = await exportLink(
      collector,
      id,
      token,
      picture(fakePng(200, 100)),
    );
    const store = await freshStore();

    const summary = await pasteLinks(collector.bundle(), store);

    expect(summary).toMatchObject({
      links: 1,
      updated: 0,
      current: 0,
      waiting: 1,
      failed: 0,
    });
    expect(summarizePaste(summary)).toBe(
      "Pasted 1 link: 1 waiting in the Inbox.",
    );
    expect(
      presentation.slides.every((slide) => slide.shapes.length === 0),
    ).toBe(true);

    const localWs = await localWorkspace();
    const inbox = await listInbox(localWs, store);
    expect(inbox.map((row) => row.id)).toEqual([item.id]);
    const placed = await insertFromInbox(inbox[0]!, localWs, store);
    expect(placed.overlapping).toBe(false);
    expect(presentation.slides.flatMap((slide) => slide.shapes)).toHaveLength(
      1,
    );
  });

  it("never downgrades: an older bundle pasted after a newer one counts as up to date", async () => {
    boot();
    const id = newId();
    const token = newToken();
    plantShape(presentation, id, token, 1);
    const store = await freshStore();

    const newer = new LocalCollector(() => 2000);
    await exportLink(newer, id, token, picture(fakePng(400, 200)), 1);
    await pasteLinks(newer.bundle(), store);
    expect(presentation.slides[1]!.shapes[0]!.fillImage).toBe(
      fakePng(400, 200),
    );

    const older = new LocalCollector(() => 1000);
    await exportLink(older, id, token, picture(fakePng(100, 50)), 1);
    const summary = await pasteLinks(older.bundle(), store);

    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    // The newer picture stays; the older bundle's was never painted.
    expect(presentation.slides[1]!.shapes[0]!.fillImage).toBe(
      fakePng(400, 200),
    );
  });

  it("fails one row's blob without stopping the rest", async () => {
    boot();
    const goodId = newId();
    const goodToken = newToken();
    const badId = newId();
    const badToken = newToken();
    plantShape(presentation, goodId, goodToken, 1);
    plantShape(presentation, badId, badToken, 1);
    const collector = new LocalCollector();
    await exportLink(
      collector,
      goodId,
      goodToken,
      picture(fakePng(400, 200)),
      1,
    );
    // Sealed with a key that is not badToken's: the shape's own token will
    // never open it.
    const wrongKeys = await deriveLinkKeys(newToken());
    await collector.putLink(
      badId,
      wrongKeys.auth,
      await seal(
        wrongKeys.enc,
        badId,
        encodePayload(picture(fakePng(400, 200))),
      ),
      1,
    );

    const store = await freshStore();
    const summary = await pasteLinks(collector.bundle(), store);

    expect(summary.updated).toBe(1);
    expect(summary.failed).toBe(1);
    expect(summary.failures).toHaveLength(1);
  });
});

describe("summarizePaste", () => {
  it("counts failures alongside the rest, comma-joined, one full stop", () => {
    expect(
      summarizePaste({
        links: 2,
        updated: 1,
        current: 0,
        waiting: 0,
        failed: 1,
        failures: [],
        notes: [],
      }),
    ).toBe("Pasted 2 links: 1 updated, 1 failed.");
  });

  it("drops the link count for an inbox-only paste", () => {
    expect(
      summarizePaste({
        links: 0,
        updated: 0,
        current: 0,
        waiting: 2,
        failed: 0,
        failures: [],
        notes: [],
      }),
    ).toBe("Pasted: 2 waiting in the Inbox.");
  });
});
