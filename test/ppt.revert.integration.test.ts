// "Revert last update" against the fake host and the fake relay: the picture
// and the tag go back one revision, a link already at its first revision has
// nowhere to go, the next update repaints the newer revision again, and a
// previous revision the relay no longer holds is a named failure. Strict load
// semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_LINK } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type * as RevertModule from "../src/ppt/revert";

enableStrictLoadSemantics();

const BEFORE = fakePng(800, 400);
const AFTER = fakePng(1600, 800);

let links: typeof LinksModule;
let revert: typeof RevertModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  revert = await import("../src/ppt/revert");
});
afterEach(() => {
  uninstallFakePpt();
});

// Insert at rev 1, push a second render, update: the deck now holds rev 2 and
// the relay still holds rev 1 - the state every test below starts from.
async function updatedLink(): Promise<{
  item: Awaited<ReturnType<typeof seedLink>>;
  rows: LinksModule.LinkRow[];
}> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedLink(BEFORE);
  await links.insertFromInbox(item, ws, relay);
  await pushAgain(item, AFTER);
  await links.updateLinks(await links.listLinks(relay), relay);
  return { item, rows: await links.listLinks(relay) };
}

function shape(): FakePptShape {
  return presentation.slides[0]!.shapes[0]!;
}

function tagRev(): number {
  const tag: unknown = JSON.parse(shape().tags.get(TAG_LINK)!);
  return (tag as { rev: number }).rev;
}

describe("revert", () => {
  it("repaints the previous revision and writes its rev back into the tag", async () => {
    const { rows } = await updatedLink();
    expect([shape().fillImage, tagRev()]).toEqual([AFTER, 2]);

    const summary = await revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 1, noPrevious: 0, failed: 0 });
    expect(summary.failures).toEqual([]);
    expect([shape().fillImage, tagRev()]).toEqual([BEFORE, 1]);
    expect(revert.summarizeRevert(summary)).toBe("1 reverted");
  });

  it("reports a link that is already at its first revision", async () => {
    const { rows } = await updatedLink();
    await revert.revertLinks(rows, relay);

    const summary = await revert.revertLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ reverted: 0, noPrevious: 1, failed: 0 });
    expect(shape().fillImage).toBe(BEFORE);
    expect(revert.summarizeRevert(summary)).toBe(
      "1 without a previous version",
    );
  });

  it("leaves the newer revision available to the next update", async () => {
    const { rows } = await updatedLink();
    await revert.revertLinks(rows, relay);

    const back = await links.listLinks(relay);
    expect(back.map((row) => row.status)).toEqual(["updateAvailable"]);
    const summary = await links.updateLinks(back, relay);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect([shape().fillImage, tagRev()]).toEqual([AFTER, 2]);
  });

  it("names the link when the relay no longer holds the previous revision", async () => {
    const { item, rows } = await updatedLink();
    // What a third push does on the server: the two-revision rule drops rev 1.
    const stored = relay.links.get(item.id)!;
    delete stored.previous;

    const summary = await revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 0, noPrevious: 0, failed: 1 });
    expect(summary.failures).toEqual([
      "revert Model!B4:F12: the relay no longer holds version 1.",
    ]);
    // Nothing was repainted, so the deck still shows what it did before.
    expect([shape().fillImage, tagRev()]).toEqual([AFTER, 2]);
    expect(revert.summarizeRevert(summary)).toBe("1 failed");
  });

  it("says so when there is nothing to revert at all", async () => {
    const summary = await revert.revertLinks([], relay);
    expect(revert.summarizeRevert(summary)).toBe("Nothing to revert");
  });
});
