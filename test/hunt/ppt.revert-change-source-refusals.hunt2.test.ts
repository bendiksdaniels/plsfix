// Hunt pass 2, target 1: failure injection at every sync step of Revert
// and Change source, refused or hung (settled with settleHungSync).
// Invariant: a plain sentence, the pane never stuck busy, and the link's
// content and tag left exactly as they were before a failed attempt.

import { afterEach, describe, expect, it, vi } from "vitest";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "../fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "../ppt.support";
import type * as ChangeSourceModule from "../../src/ppt/change-source";
import type * as LinksModule from "../../src/ppt/links";
import type * as RevertModule from "../../src/ppt/revert";
import { settleHungSync } from "../hung-sync";

enableStrictLoadSemantics();

function expectPlainSentence(message: string): void {
  expect(message).not.toMatch(/InvalidArgument|ItemNotFound|GeneralException/);
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

interface Modules {
  links: typeof LinksModule;
  changeSource: typeof ChangeSourceModule;
  revert: typeof RevertModule;
}

async function importModules(): Promise<Modules> {
  return {
    links: await import("../../src/ppt/links"),
    changeSource: await import("../../src/ppt/change-source"),
    revert: await import("../../src/ppt/revert"),
  };
}

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

// ---------------------------------------------------------------------------
// Revert
// ---------------------------------------------------------------------------

async function revertScenario(): Promise<{
  helpers: FakePptHelpers;
  presentation: FakePresentation;
  relay: FakeRelay;
  modules: Modules;
  rows: Awaited<ReturnType<Modules["links"]["listLinks"]>>;
}> {
  const { helpers, presentation, relay } = await bootPpt();
  const modules = await importModules();
  const ws: Workspace = await createWorkspace(memoryStore());
  const item = await seedLink(fakePng(200, 100));
  await modules.links.insertFromInbox(item, ws, relay);
  await pushAgain(item, fakePng(400, 200));
  await modules.links.updateLinks(await modules.links.listLinks(relay), relay);
  const rows = await modules.links.listLinks(relay);
  return { helpers, presentation, relay, modules, rows };
}

// A picture's repaint (applyBatch -> paintBatch, links.ts) is not one sync:
// it tries the whole-deck BATCH route first and, on anything but a draw
// timeout, silently retries the SAME row one at a time through refreshLink -
// deliberate resilience (links.ts's own comment: "those rows go through one
// at a time, only the bad one is reported"), not a bug. So "the Nth sync"
// here means the Nth of the up-to-two ATTEMPTS a one-row revert can spend,
// not a fixed sync count: refusing attempt 0 alone is absorbed by attempt 1
// and still succeeds (pinned below as its own case, not a failure); only
// refusing BOTH exhausts the retry into a genuine, reported failure.
const REVERT_ATTEMPTS = 2;

describe("Revert: a single transient refusal is absorbed by the retry", () => {
  it("attempt 0 refused, attempt 1 unarmed: still reverts, no failure reported", async () => {
    const { helpers, relay, modules, rows } = await revertScenario();
    helpers.refuseNextSync(new Error("the host answered with an error"), 0);

    const summary = await modules.revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 1, failed: 0 });
  });
});

describe(`Revert: every attempt refused (${String(REVERT_ATTEMPTS)} in a row), a genuine failure`, () => {
  // Not asserted here: that the picture and tag stay at the CURRENT
  // revision. queueRefresh (refresh.ts) writes fill.setImage, height and
  // the tag as plain property writes on the shape that already exists, and
  // test/fakeppt/objects.ts's setters apply every one of them straight to
  // the model with no pending/rollback tracking at all (only a shape ADD is
  // tracked, in FakePresentation.pendingIds) - proven directly: even a HUNG
  // sync's rollbackPending() leaves the picture repainted, because nothing
  // it does touches a plain property write. Reported under OUT-OF-SLICE
  // (test/fakeppt/model.ts); the three writes travel together in one
  // queueRefresh call regardless, so even where the fake's shortcut lets a
  // "failed" attempt still land, it lands as one coherent picture+tag pair,
  // never a split between them - which is what the fixes elsewhere in this
  // hunt (cleanupByToken, change-source's rollback) actually guard against.
  it("fails the row with a plain sentence, single shape, no crash", async () => {
    const { helpers, presentation, relay, modules, rows } =
      await revertScenario();
    for (let n = 0; n < REVERT_ATTEMPTS; n += 1) {
      helpers.refuseNextSync(new Error("the host answered with an error"), n);
    }

    const summary = await modules.revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 0, failed: 1 });
    expectPlainSentence(summary.failures[0]!);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });

  it("a hung first attempt settles (never stuck), reports it, never retrying into a double paint", async () => {
    // isDrawTimeout's own branch in paintBatch (links.ts): a hang becomes a
    // ChartDrawTimeout at the deadline, which fails every row in the batch
    // directly - the retry-below is for anything else, deliberately never
    // for a host that has genuinely stopped answering (its own comment: "a
    // timeout fails every row here instead ... leaves the next batch to try
    // its own round trip fresh"). So a hang never reaches a second attempt.
    const { helpers, presentation, relay, modules, rows } =
      await revertScenario();
    helpers.hangNextSync(0);
    vi.useFakeTimers();

    const summary = await settleHungSync(
      modules.revert.revertLinks(rows, relay),
    );

    expect(summary).toMatchObject({ reverted: 0, failed: 1 });
    expectPlainSentence(summary.failures[0]!);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Change source
// ---------------------------------------------------------------------------

async function changeSourceScenario(): Promise<{
  helpers: FakePptHelpers;
  presentation: FakePresentation;
  relay: FakeRelay;
  modules: Modules;
  row: LinksModule.LinkRow;
  newer: Awaited<ReturnType<typeof seedLink>>;
  ws: Workspace;
}> {
  const { helpers, presentation, relay } = await bootPpt();
  const modules = await importModules();
  const ws: Workspace = await createWorkspace(memoryStore());
  const item = await seedLink(fakePng(200, 100));
  await modules.links.insertFromInbox(item, ws, relay);
  const row = (await modules.links.listLinks(relay))[0]!;
  const newer = await seedLink(fakePng(400, 200), "Model_v5.xlsx");
  return { helpers, presentation, relay, modules, row, newer, ws };
}

// changeSource has two stages: retagLink (its own single sync, no retry at
// all) then the repaint, which - like revert above - is the same
// batch-then-per-row-retry path, so a lone refusal at the repaint's first
// attempt is absorbed exactly as it is there. "Sync 0" below is always
// retagLink's own; "sync 1" is the repaint's first attempt, which succeeds
// on its own retry unless refused twice in a row.
describe("Change source: retagLink's own sync refused (sync 0), no retry for this stage", () => {
  it("rejects cleanly and rolls back to the OLD identity - the fix: a host may apply the tag write before answering the sync with an error", async () => {
    // Pins the exact gap change-source.ts's own comment used to get wrong:
    // "Nothing was written yet, so there is nothing to roll back" assumed a
    // refused sync means nothing landed - the same Office.js caveat this
    // hunt already found for chart-draw.ts's grouping sync says otherwise.
    const { helpers, presentation, relay, modules, row, newer, ws } =
      await changeSourceScenario();
    const oldTag = presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK");
    helpers.refuseNextSync(new Error("the host answered with an error"), 0);

    await expect(
      modules.changeSource.changeSource(row, newer, ws, relay),
    ).rejects.toThrow();

    const afterTag = presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK");
    expect(afterTag).toBe(oldTag);
  });

  it("a hung retagLink settles, never stuck, and leaves the OLD identity in place", async () => {
    const { helpers, presentation, relay, modules, row, newer, ws } =
      await changeSourceScenario();
    const oldTag = presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK");
    helpers.hangNextSync(0);
    vi.useFakeTimers();

    await expect(
      settleHungSync(modules.changeSource.changeSource(row, newer, ws, relay)),
    ).rejects.toThrow();

    expect(presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK")).toBe(
      oldTag,
    );
  });
});

describe("Change source: the repaint's lone first attempt refused (sync 1), absorbed by its own retry", () => {
  it("still succeeds - the SAME resilience revert relies on, exercised here too", async () => {
    const { helpers, relay, modules, row, newer, ws } =
      await changeSourceScenario();
    helpers.refuseNextSync(new Error("the host answered with an error"), 1);

    const summary = await modules.changeSource.changeSource(
      row,
      newer,
      ws,
      relay,
    );

    expect(summary).toContain("-> Model_v5.xlsx");
  });
});

describe("Change source: the repaint refused on both its attempts, a genuine failure", () => {
  // The tag is checked, not the picture: rollback() only ever retags (it
  // calls host.retagLink, never a repaint), so if queueRefresh's own
  // immediate writes (see the OUT-OF-SLICE note above Revert) had already
  // painted the NEW picture before both attempts were reported failed, the
  // rollback below would still leave the OLD tag over a picture that had
  // already moved on - the same fake gap, arrived at from the other
  // direction. Not reachable through this suite's own assertions today (no
  // test here reads fillImage in this exact case), so left named rather
  // than silently assumed away.
  it("rejects cleanly and rolls back to the OLD identity, never a new tag over the old picture", async () => {
    const { helpers, presentation, relay, modules, row, newer, ws } =
      await changeSourceScenario();
    const oldTag = presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK");
    helpers.refuseNextSync(new Error("the host answered with an error"), 1);
    helpers.refuseNextSync(new Error("the host answered with an error"), 2);

    let threw: unknown;
    try {
      await modules.changeSource.changeSource(row, newer, ws, relay);
    } catch (error) {
      threw = error;
    }

    expect(threw).toBeDefined();
    expectPlainSentence(messageOf(threw));
    const afterTag = presentation.slides[0]!.shapes[0]!.tags.get("PLSFIX_LINK");
    expect(afterTag).toBe(oldTag);
  });
});
