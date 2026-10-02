// Hunt pass 2: cleanupByToken's own token is per LINK, not per push, so a
// second insert of a link already on the slide must never delete the user's
// earlier, finished copy - and, since the cleanup can hang exactly like the
// add it follows, must never leave the pane stuck either.

import { afterEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../../src/link/model";
import { createWorkspace } from "../../src/link/workspace";
import { fakePng } from "../fakepng";
import { enableStrictLoadSemantics, uninstallFakePpt } from "../fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "../ppt.support";
import { settleHungSync } from "../hung-sync";

enableStrictLoadSemantics();

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

type Seed = () => Promise<InboxItem>;

// The third element is a fixed add-sync index, used instead of measuring a
// clean second insert: only table needs it, because its own add+tag+load
// sync is followed by formatNewTable's format syncs, so "the last sync of a
// clean run" would land on a format sync instead of the add sync this pin
// means to fail. 2 is resolveTarget's own two read-only syncs (the selected
// slide, then the slide's shapes - a pre-existing Table triggers neither the
// frame-look nor the text-extent read the other kinds' own shapes do).
const KINDS: [string, Seed, number?][] = [
  ["picture", () => seedLink(fakePng(200, 100))],
  ["text", () => seedText("EUR 15.7m")],
  ["table", () => seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]), 2],
  ["chart", () => seedChart(columnChart(2), fakePng(400, 200))],
];

// Boots, inserts the item once (the user's legit shape), and returns what a
// SECOND insert of the same item onto the same slide needs.
async function withLegitShape(seed: Seed) {
  const booted = await bootPpt();
  const ws = await createWorkspace(memoryStore());
  const item = await seed();
  await booted.links.insertFromInbox(item, ws, booted.relay);
  const legit = booted.presentation.slides[0]!.shapes[0]!;
  return { ...booted, ws, item, legitId: legit.id };
}

// The second insert's own sync count, measured clean: its add+tag sync (or,
// for a chart, its grouping sync) is the last one - true whenever nothing
// runs after it, which KINDS's fixed index covers for the one kind it is not.
async function secondInsertSyncs(seed: Seed): Promise<number> {
  const b = await withLegitShape(seed);
  const before = b.helpers.syncCount();
  await b.links.insertFromInbox(b.item, b.ws, b.relay);
  const total = b.helpers.syncCount() - before;
  uninstallFakePpt();
  return total;
}

describe("cleanupByToken never deletes a shape that was already there", () => {
  for (const [kind, seed, fixedIndex] of KINDS) {
    it(`${kind}: rolled back (failNextSync) - the legit shape survives, no orphan beside it`, async () => {
      const at = fixedIndex ?? (await secondInsertSyncs(seed)) - 1;
      const b = await withLegitShape(seed);
      b.helpers.failNextSync(new Error("the host answered with an error"), at);

      await expect(
        b.links.insertFromInbox(b.item, b.ws, b.relay),
      ).rejects.toThrow();

      const ids = b.presentation.slides[0]!.shapes.map((s) => s.id);
      expect(ids).toEqual([b.legitId]);
    });

    it(`${kind}: applied then refused (refuseNextSync) - the legit shape survives, no orphan beside it`, async () => {
      const at = fixedIndex ?? (await secondInsertSyncs(seed)) - 1;
      const b = await withLegitShape(seed);
      b.helpers.refuseNextSync(
        new Error("the host answered with an error"),
        at,
      );

      await expect(
        b.links.insertFromInbox(b.item, b.ws, b.relay),
      ).rejects.toThrow();

      const ids = b.presentation.slides[0]!.shapes.map((s) => s.id);
      expect(ids).toEqual([b.legitId]);
    });
  }
});

// picture, text and table: the three kinds whose insert calls cleanupByToken
// directly (host.ts, texts.ts, tables.ts) with no deadline of its own before
// this fix. Chart is not repeated here: chart-draw.ts's cleanupWithin already
// ran its cleanupByToken call under withSyncDeadline before this hunt.
const STUCK_CASES: [string, Seed, string][] = [
  ["picture", () => seedLink(fakePng(200, 100)), "inserting the picture"],
  ["text", () => seedText("EUR 15.7m"), "inserting the text"],
  [
    "table",
    () => seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]),
    "inserting the table",
  ],
];

async function freshInsert(seed: Seed) {
  const booted = await bootPpt();
  const ws = await createWorkspace(memoryStore());
  const item = await seed();
  return { ...booted, ws, item };
}

// Sync order on a fresh, empty slide, for all three kinds alike: 0 the
// selected slide, 1 the slide's shapes (both read-only, "free" placement),
// 2 the kind's own add + both tags + load(id) - table's later format syncs
// (formatNewTable) never run because the add sync itself is what hangs, so
// they cannot be mistaken for it the way measuring a full clean run would.
const ADD_SYNC_INDEX = 2;

describe("a hung add sync never leaves the pane stuck, even when the cleanup it triggers hangs too", () => {
  for (const [kind, seed, what] of STUCK_CASES) {
    it(`${kind}: settles and rejects with the add's own timeout sentence`, async () => {
      const b = await freshInsert(seed);
      // The add+tag sync hangs, and (lessons: "every later batch queues
      // behind it") so does cleanupByToken's own first sync right after it.
      b.helpers.hangNextSync(ADD_SYNC_INDEX);
      b.helpers.hangNextSync(ADD_SYNC_INDEX + 1);
      vi.useFakeTimers();

      let threw: unknown;
      try {
        await settleHungSync(b.links.insertFromInbox(b.item, b.ws, b.relay));
      } catch (error) {
        threw = error;
      }

      expect(threw).toBeInstanceOf(Error);
      expect((threw as Error).message).toBe(
        `PowerPoint stopped answering while ${what}`,
      );
    });
  }
});
