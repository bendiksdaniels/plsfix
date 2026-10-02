// Hunt pass 2, target 1: failure injection at EVERY sync step of every
// insert kind (picture, table, text, chart), generated over the step index -
// for N = 0..last, the Nth PowerPoint.run sync of the insert is made to
// refuse (refuseNextSync) or hang (hangNextSync, settled with
// settleHungSync). Invariant after each: a plain sentence (no raw Office
// text), the pane not left stuck busy, and no untagged or half-built shape
// left on the slide.
//
// The sweep found a real gap common to every single-sync insert (a plain
// picture, a text box, or a table before its own format chunks run): the
// add and both tags queue in the SAME batch as the one sync that confirms
// them, and a real host may apply that batch and still answer the sync()
// call itself with an error (drawGroup's own comment already names this
// Office.js caveat for a chart's many syncs - "the batch's adds stay on the
// deck, only the loads are lost" - it just never applied it to the ONE sync
// every other insert kind has). That left a fully tagged, orphaned shape
// nothing ever removed, because its id was never read back to clean up by.
// Fixed at the root: chart-cleanup.ts's new cleanupByToken finds it by the
// one tag no two links ever share instead, wired into tables.ts, texts.ts,
// host.ts's insertPictureInPlace and chart-picture.ts's pictureSynced.

import { afterEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../../src/link/model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "../fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "../ppt.support";
import type * as LinksModule from "../../src/ppt/links";
import { settleHungSync } from "../hung-sync";

enableStrictLoadSemantics();

// No raw Office code ever reaches the toast - the sentinel every hunt file
// in this folder checks a plain-sentence error against.
function expectPlainSentence(message: string): void {
  expect(message).not.toMatch(/InvalidArgument|ItemNotFound|GeneralException/);
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

interface Scenario {
  name: string;
  seed: () => Promise<InboxItem>;
}

const SCENARIOS: Scenario[] = [
  { name: "picture", seed: () => seedLink(fakePng(200, 100)) },
  {
    name: "table",
    seed: () => seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]),
  },
  { name: "text", seed: () => seedText("EUR 15.7m") },
  { name: "chart", seed: () => seedChart(columnChart(2), fakePng(400, 200)) },
];

interface Booted {
  links: typeof LinksModule;
  presentation: FakePresentation;
  helpers: FakePptHelpers;
  relay: FakeRelay;
  ws: Workspace;
  item: InboxItem;
}

// One fresh boot + seeded inbox item per call, so every sweep iteration (and
// the one clean measurement run) starts from identical, untouched state.
async function freshInsert(seed: Scenario["seed"]): Promise<Booted> {
  const { links, presentation, helpers, relay } = await bootPpt();
  const ws = await createWorkspace(memoryStore());
  const item = await seed();
  return { links, presentation, helpers, relay, ws, item };
}

function insertOf(booted: Booted): Promise<unknown> {
  return booted.links.insertFromInbox(booted.item, booted.ws, booted.relay);
}

// Measured once per kind, clean, before any test registers: the exact
// number of PowerPoint.run syncs a successful insert of that kind spends,
// so the sweep below covers N = 0..last precisely instead of a guess.
async function measureTotalSyncs(seed: Scenario["seed"]): Promise<number> {
  const booted = await freshInsert(seed);
  const before = booted.helpers.syncCount();
  await insertOf(booted);
  const total = booted.helpers.syncCount() - before;
  uninstallFakePpt();
  return total;
}

// Sequential, deliberately: bootPpt() installs the fake globals afresh each
// call, so running these measurements concurrently (Promise.all) would have
// one scenario's uninstall/install race another's still-in-flight insert.
const TOTALS = new Map<string, number>();
for (const scenario of SCENARIOS) {
  TOTALS.set(scenario.name, await measureTotalSyncs(scenario.seed));
}

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

describe("insert: every sync step refused, one case per step", () => {
  for (const scenario of SCENARIOS) {
    const total = TOTALS.get(scenario.name)!;
    describe(`${scenario.name} (${String(total)} syncs when it succeeds)`, () => {
      for (let n = 0; n < total; n += 1) {
        it(`refusing sync ${String(n)} leaves no shape and a plain sentence`, async () => {
          const booted = await freshInsert(scenario.seed);
          booted.helpers.refuseNextSync(
            new Error("PowerPoint could not complete the request."),
            n,
          );

          let threw: unknown;
          try {
            await insertOf(booted);
          } catch (error) {
            threw = error;
          }

          expect(threw).toBeDefined();
          expectPlainSentence(messageOf(threw));
          expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
        });
      }
    });
  }
});

describe("insert: every sync step hung, one case per step", () => {
  for (const scenario of SCENARIOS) {
    const total = TOTALS.get(scenario.name)!;
    describe(`${scenario.name} (${String(total)} syncs when it succeeds)`, () => {
      for (let n = 0; n < total; n += 1) {
        it(`a hung sync ${String(n)} settles (never stuck) and leaves no stray shape`, async () => {
          const booted = await freshInsert(scenario.seed);
          booted.helpers.hangNextSync(n);
          vi.useFakeTimers();

          let threw: unknown;
          let settled = false;
          try {
            await settleHungSync(insertOf(booted));
            settled = true;
          } catch (error) {
            threw = error;
            settled = true;
          }

          // The whole point: the promise settles at all (rejects or
          // resolves), never left hanging past the sync deadline.
          expect(settled).toBe(true);

          if (threw !== undefined) {
            expectPlainSentence(messageOf(threw));
            expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
          } else {
            // Resolved: either a normal insert (the hang was armed past the
            // real total and never fired) or - chart only - the timeout's
            // own picture fallback. Either way exactly one finished, tagged
            // shape, never a stray beside a half-built one.
            expect(booted.presentation.slides[0]!.shapes.length).toBe(1);
          }
        });
      }
    });
  }
});

// The exact regression the sweep above found: the single sync that adds a
// shape and both its tags, refused with the host having applied it first -
// the one case cleanupByToken exists for. Pinned on its own, deliberately,
// so this specific mechanism has a named regression test independent of the
// generic sweep above (which covers it too, among dozens of other steps).
describe("cleanupByToken: the add+tag sync itself refused, host applied it first", () => {
  it("table: no orphaned, half-formatted table left on the slide", async () => {
    const booted = await freshInsert(
      SCENARIOS.find((s) => s.name === "table")!.seed,
    );
    // Sync order for this scenario: 1 the selected slide, 2 the slide's
    // shapes (both read-only), 3 the table's own add + both tags + load(id).
    booted.helpers.refuseNextSync(
      new Error("the host answered the sync with an error"),
      2,
    );

    await expect(insertOf(booted)).rejects.toThrow();

    expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("text: no orphaned text box left on the slide", async () => {
    const booted = await freshInsert(
      SCENARIOS.find((s) => s.name === "text")!.seed,
    );
    booted.helpers.refuseNextSync(
      new Error("the host answered the sync with an error"),
      2,
    );

    await expect(insertOf(booted)).rejects.toThrow();

    expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("picture: no orphaned picture left on the slide", async () => {
    const booted = await freshInsert(
      SCENARIOS.find((s) => s.name === "picture")!.seed,
    );
    // The picture's own add + both tags + load(id) is the LAST of its
    // syncs (TOTALS.get("picture") - 1), not sync 1: that one is only the
    // free-space read, still occupied by nothing at all, so refusing it is
    // a no-op refusal that leaves no shape regardless of cleanupByToken.
    booted.helpers.refuseNextSync(
      new Error("the host answered the sync with an error"),
      TOTALS.get("picture")! - 1,
    );

    await expect(insertOf(booted)).rejects.toThrow();

    expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("a second, later insert of the same export lands cleanly after the first was cleaned up", async () => {
    // Proves cleanupByToken's own promise, not merely that no shape is left
    // in THIS run: a retry never lands beside a phantom from the failed one.
    const booted = await freshInsert(
      SCENARIOS.find((s) => s.name === "table")!.seed,
    );
    booted.helpers.refuseNextSync(
      new Error("the host answered the sync with an error"),
      2,
    );
    await expect(insertOf(booted)).rejects.toThrow();
    expect(booted.presentation.slides[0]!.shapes).toHaveLength(0);

    await insertOf(booted);

    expect(booted.presentation.slides[0]!.shapes).toHaveLength(1);
  });
});

// The boundary the fix above had to get right the second time: a REFRESH's
// grouping sync carries the SAME token as the old group it is about to
// replace (a link keeps its identity across updates), so cleanupByToken must
// never run there - only insertChart, never refreshChart, passes no `before`
// to drawGroup, and that is the one signal cleanupWithin trusts.
//
// The rolled-back (failNextSync) half of this boundary is NOT pinned here:
// test/ppt.charts.audit.integration.test.ts's own "a refresh the host
// refuses outright" already proves the old group survives a rolled-back
// FIRST-chunk failure, and reproducing the SAME rollback at the LATER
// grouping sync ran into a fake limitation reported under OUT-OF-SLICE
// (test/fakeppt/model.ts: shape.delete() is never tracked as pending, so it
// is never undone by rollbackPending() the way an add is - a delete queued
// in the same batch as the grouping sync takes effect immediately and stays
// gone even when that sync is rolled back). That gap belongs to the fake,
// not to this fix, and is not safe to change this late without auditing
// every existing test that calls delete() near a rollback.
describe("cleanupByToken must never touch a refresh's still-valid old group", () => {
  it("the grouping sync refused with the batch already applied (refuseNextSync) leaves one correctly tagged group, no orphan beside it", async () => {
    // The other half of the same boundary: applied:true means old.delete()
    // landed too, so the deck's one shape afterwards is the NEW group - not
    // an orphan beside a surviving old one, and not a duplicate. Skipping
    // cleanupByToken here costs nothing: there is nothing left to clean up.
    const booted = await freshInsert(
      SCENARIOS.find((s) => s.name === "chart")!.seed,
    );
    await insertOf(booted);
    const oldShapeId = booted.presentation.slides[0]!.shapes[0]!.id;
    await pushChart(booted.item, columnChart(3), fakePng(400, 200));
    const rows = await booted.links.listLinks(booted.relay);

    booted.helpers.refuseNextSync(
      new Error("the host answered the sync with an error"),
      1,
    );
    const summary = await booted.links.updateLinks(rows, booted.relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    const shapesNow = booted.presentation.slides[0]!.shapes;
    expect(shapesNow).toHaveLength(1);
    expect(shapesNow[0]!.id).not.toBe(oldShapeId);
    expect(shapesNow[0]!.tags.has("PLSFIX_LINK")).toBe(true);
    expect(shapesNow[0]!.tags.has("PLSFIX_KEY")).toBe(true);
  });
});
