// J audit, part two: the repaint flows built on links.ts applyBatch - "Update
// all", "Revert last update" and "Change source" - all share one batched
// repaint, so a host that swallows its sync must fail every row of that batch
// with the sentence and let the update end, never retry the same hang row by
// row. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { SYNC_TIMEOUT_MS } from "../src/ppt/chart-draw";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "./ppt.support";
import type * as ChangeSourceModule from "../src/ppt/change-source";
import type * as LinksModule from "../src/ppt/links";
import type * as RevertModule from "../src/ppt/revert";

enableStrictLoadSemantics();

const BEFORE = fakePng(800, 400);
const AFTER = fakePng(1600, 800);

let links: typeof LinksModule;
let revert: typeof RevertModule;
let changeSource: typeof ChangeSourceModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  revert = await import("../src/ppt/revert");
  changeSource = await import("../src/ppt/change-source");
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

function shape(index = 0): FakePptShape {
  return presentation.slides[0]!.shapes[index]!;
}

// See j.audit.integration.test.ts: the same clock-driven settle, copied
// rather than shared because each audit file owns its own suite.
async function settle<T>(work: Promise<T>): Promise<T> {
  let done = false;
  void work.then(
    () => (done = true),
    () => (done = true),
  );
  for (let step = 0; step < 10 && !done; step += 1) {
    await vi.advanceTimersByTimeAsync(SYNC_TIMEOUT_MS);
  }
  return work;
}

describe('"Update all" over a batch the host never answers', () => {
  it("fails every row of the swallowed batch, and the update ends there", async () => {
    // The same link on two slides - "a deck can hold the same link on twenty
    // slides", per keyLinks's own comment - so the batch holds two rows from
    // one token and one real key derivation, not two.
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(BEFORE);
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, AFTER);
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();

    // Two plain pictures repaint in one host batch, so one hang catches both.
    helpers.hangNextSync();
    vi.useFakeTimers();
    const summary = await settle(links.updateLinks(rows, relay));

    expect(summary).toMatchObject({ updated: 0, failed: 2 });
    expect(summary.failures).toHaveLength(2);
    for (const line of summary.failures) {
      expect(line).toMatch(/PowerPoint stopped answering while repainting/);
    }
    // Exactly the one attempt: the old per-row fallback would have spent two
    // more syncs retrying a host that had genuinely stopped answering - and,
    // unarmed to hang again, would have quietly reported both as updated.
    expect(helpers.syncCount()).toBe(before + 1);

    // The update ended - the pane is not stuck busy - and the next one is not
    // jammed by the one that hung: it finds both rows already painted (the
    // fake, like the real host, cannot tell the caller whether a batch it
    // never confirmed actually landed; queueRefresh's writes happen before
    // the sync that swallowed them) and reports nothing left to do.
    vi.useRealTimers();
    const again = await links.updateLinks(await links.listLinks(relay), relay);
    expect(again).toMatchObject({ updated: 0, current: 2, failed: 0 });
  });
});

describe("a revert the host never answers", () => {
  // Insert at rev 1, push a second render, update: the deck holds rev 2 and
  // the relay still holds rev 1, the state every revert starts from (mirrors
  // ppt.revert.integration.test.ts's own helper).
  async function updatedLink(): Promise<{
    item: InboxItem;
    rows: LinksModule.LinkRow[];
  }> {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(BEFORE);
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, AFTER);
    await links.updateLinks(await links.listLinks(relay), relay);
    return { item, rows: await links.listLinks(relay) };
  }

  it("names what stopped instead of hanging, and a later revert still lands", async () => {
    const { rows } = await updatedLink();
    helpers.hangNextSync();
    vi.useFakeTimers();

    const summary = await settle(revert.revertLinks(rows, relay));

    expect(summary).toMatchObject({ reverted: 0, noPrevious: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(/stopped answering/);

    // The revert reported a failure, not a hang, and the pane can still act:
    // an "Update all" right afterwards is a fresh round trip, not the one
    // that never came back, and it lands normally.
    vi.useRealTimers();
    const again = await links.updateLinks(await links.listLinks(relay), relay);
    expect(again).toMatchObject({ updated: 1, failed: 0 });
    expect(shape().fillImage).toBe(AFTER);
  });
});

describe("a change of source the host never answers", () => {
  async function seededRow(): Promise<{
    ws: Workspace;
    row: LinksModule.LinkRow;
    newer: InboxItem;
  }> {
    const ws = await createWorkspace(memoryStore());
    const placed = await seedLink(BEFORE);
    await links.insertFromInbox(placed, ws, relay);
    const newer = await seedLink(AFTER, "Model_v5.xlsx");
    const row = (await links.listLinks(relay))[0]!;
    return { ws, row, newer };
  }

  it("puts the old tag back and names what stopped when the repaint hangs", async () => {
    const { ws, row, newer } = await seededRow();
    const before = shape().tags.get(TAG_LINK);
    // The retag lands (its own sync answers); the repaint that follows it -
    // the shared batch path, one row - is the one that never comes back.
    helpers.hangNextSync(1);
    vi.useFakeTimers();

    await expect(
      settle(changeSource.changeSource(row, newer, ws, relay)),
    ).rejects.toThrow(/PowerPoint stopped answering while repainting/);

    // The rollback's own retag is a normal sync and still lands, so the shape
    // claims its old link again - what a swallowed batch actually left on the
    // slide is exactly what nobody can tell, on the fake or the real host,
    // which is why the tag identity, not the picture, is what gets put back.
    expect(shape().tags.get(TAG_LINK)).toBe(before);

    // The item never left the inbox, so the same change of source can be
    // tried again once the host answers.
    vi.useRealTimers();
    const summary = await changeSource.changeSource(row, newer, ws, relay);
    expect(summary).toBe("Source changed: Model_v4.xlsx -> Model_v5.xlsx");
    expect(shape().fillImage).toBe(AFTER);
  });
});
