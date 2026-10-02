// Hunt pass 1: small uncovered branches across the ppt adapter -
// overlapNote's "the largest free spot is W x H pt" sentence (never called
// with a freeSpotSize directly), and a text link refresh failing for a
// reason OTHER than a deleted shape (missing-shape is proven elsewhere).

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createWorkspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import { enableStrictLoadSemantics, uninstallFakePpt } from "../fakeppt";
import { bootPpt, memoryStore, pushText, seedText } from "../ppt.support";
import { overlapNote, OVERLAP_NOTE } from "../../src/ppt/host";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

describe("overlapNote: the largest-free-spot sentence", () => {
  it("names the size once free-space scanning found a hole, no spot label needed", () => {
    expect(overlapNote(undefined, { width: 888, height: 40 })).toBe(
      "Placed over other objects: the largest free spot is 888 x 40 pt",
    );
  });

  it("the size wins over a spot label when both are given (a Selected-shape overlap with a scan behind it)", () => {
    expect(overlapNote("bottom-right", { width: 10, height: 320 })).toBe(
      "Placed over other objects: the largest free spot is 10 x 320 pt",
    );
  });

  it("still falls back to the plain sentence with neither", () => {
    expect(overlapNote()).toBe(OVERLAP_NOTE);
  });
});

let links: typeof LinksModule;
let relay: FakeRelay;
let helpers: Awaited<ReturnType<typeof bootPpt>>["helpers"];

beforeEach(async () => {
  ({ links, relay, helpers } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

describe("a text link refresh refused for a reason that is not a missing shape", () => {
  it("wraps the host's own message behind the link's label, same as every other adapter", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedText("EUR 15.7m");
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 16.1m");
    const rows = await links.listLinks(relay);
    helpers.failNextSync(new Error("PowerPoint stopped answering"));

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    // texts.ts's own catch already prefixes "refresh <label>: "; failureLine
    // sees its own label inside the message and passes it through unchanged.
    expect(summary.failures[0]).toBe(
      "refresh Model!B4:F12 text: PowerPoint stopped answering",
    );
  });
});
