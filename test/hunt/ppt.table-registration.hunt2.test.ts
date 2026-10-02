// Hunt pass 2: PowerPoint for the web registers a new table ~200 ms after
// the add's own sync already answered its id (rig 27.09); shapes.getItem(id)
// and any re-anchored proxy fail 5010 in that window. Pins shape-ready.ts's
// poll and chart-cleanup.ts's catch-up, via the opt-in register-delay.ts fake.

import { afterEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK, TAG_PAINT } from "../../src/link/model";
import { createWorkspace } from "../../src/link/workspace";
import { cleanupShapes } from "../../src/ppt/chart-cleanup";
import { enableStrictLoadSemantics, uninstallFakePpt } from "../fakeppt";
import { settleHungSync } from "../hung-sync";
import { bootPpt, memoryStore, pushTable, seedTable } from "../ppt.support";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

// The registration poll's own wait is real by default (100 ms); every test
// here swaps it for an instant resolve so a run that goes to the cap costs no
// real time. Dynamic, and called AFTER bootPpt: bootPpt's vi.resetModules()
// gives tables.ts and chart-cleanup.ts a fresh copy of shape-ready.ts each
// test, so the swap has to land on that same fresh copy, not an earlier one.
async function instantWait(): Promise<void> {
  const shapeReady = await import("../../src/ppt/shape-ready");
  shapeReady.setRegisterPollWait(() => Promise.resolve());
}

describe("a table insert or rebuild under the web's own registration delay", () => {
  it("insert: a table under a 3-sync delay still lands, formatted and tagged", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    await instantWait();
    const ws = await createWorkspace(memoryStore());
    helpers.delayRegistration(3);
    const item = await seedTable(
      [
        [{ t: "Q1", b: true }, { t: "Q2" }],
        [{ t: "100", a: "r" }, { t: "120" }],
      ],
      [80, 60],
    );

    // On the current code (before shape-ready.ts's poll), this fails with
    // "InvalidParam passed to GetItem(id)": formatNewTable's shape.getTable()
    // runs the instant the add's sync answers, straight into the window.
    await links.insertFromInbox(item, ws, relay);

    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(1);
    const shape = shapes[0]!;
    expect(shape.type).toBe("Table");
    const grid = shape.table!;
    expect(grid.cells.map((row) => row.map((cell) => cell.text))).toEqual([
      ["Q1", "Q2"],
      ["100", "120"],
    ]);
    expect(grid.cells[0]![0]!.font.bold).toBe(true);
    expect(grid.cells[1]![0]!.horizontalAlignment).toBe("Right");
    expect(shape.tags.get(TAG_LINK)).toBeDefined();
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(shape.tags.has(TAG_PAINT)).toBe(true);
  });

  it("rebuild: a source that grows under a 3-sync delay still rebuilds, formatted and tagged", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    await instantWait();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]);
    await links.insertFromInbox(item, ws, relay);

    helpers.delayRegistration(3);
    await pushTable(
      item,
      [
        [{ t: "Q1" }, { t: "Q2" }],
        [{ t: "100", b: true }, { t: "120" }],
      ],
      [80, 80],
    );

    // On the current code, built.getTable() in rebuildTable fails the same
    // way formatNewTable's does: the same window, one sync later.
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(1);
    const grid = shapes[0]!.table!;
    expect(grid.rowCount).toBe(2);
    expect(grid.cells[1]!.map((cell) => cell.text)).toEqual(["100", "120"]);
    expect(grid.cells[1]![0]!.font.bold).toBe(true);
    expect(JSON.parse(shapes[0]!.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      rev: 2,
    });
  });

  it("insert: a delay past the poll cap gives the pane sentence, not a raw office error", async () => {
    const { links, presentation, helpers, relay } = await bootPpt();
    await instantWait();
    const shapeReady = await import("../../src/ppt/shape-ready");
    const ws = await createWorkspace(memoryStore());
    // Comfortably past REGISTER_POLL_TRIES (30): the poll must exhaust its
    // cap and give up, never spin forever waiting for a host that (in this
    // simulation) never catches up in time.
    helpers.delayRegistration(shapeReady.REGISTER_POLL_TRIES + 5);
    const item = await seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]);

    const before = helpers.syncCount();
    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
      "PowerPoint did not finish creating the table. Try again.",
    );
    const spent = helpers.syncCount() - before;

    // Bounded: the poll's own 30 tries plus cleanupShapes's own catch-up
    // tries, never unbounded and never a raw "InvalidArgument"/"5010" string.
    expect(spent).toBeGreaterThan(shapeReady.REGISTER_POLL_TRIES);
    expect(spent).toBeLessThan(100);
    // No half-formatted orphan left behind: formatNewTable's own catch still
    // runs cleanupShapes on this failure, same as any other.
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("insert: an ordinary table with no delay spends exactly one extra sync", async () => {
    const { links, helpers, relay } = await bootPpt();
    await instantWait();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]);

    const before = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    const spent = helpers.syncCount() - before;

    // 0 the selected slide, 1 the slide's shapes (resolveTarget, both
    // read-only), 2 the add + both tags + load(id), 3 the registration poll
    // (answers not-null at once, undelayed - shape-ready.ts's whole point),
    // 4 the format step: neither cell here carries a format, so
    // writeCellsInChunks's chunk list is empty and this is the tag-only sync
    // it spends instead (its own early-return branch). Was 4 before this
    // fix: test/hunt/ppt.cleanup-token.hunt2.test.ts fixes this exact
    // scenario's add-sync at index 2, still true - the poll only ever adds
    // sync 3, never moves the ones before it.
    expect(spent).toBe(5);
  });
});

describe("cleanupShapes catches up on an id that is not yet registered", () => {
  it("deletes a confirmed id once it registers a couple of syncs later, no orphan left", async () => {
    const { presentation, helpers } = await bootPpt();
    await instantWait();
    const slideId = presentation.slides[0]!.id;
    helpers.delayRegistration(2);

    let id = "";
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      const shape = shapes.addTable(1, 1, {
        left: 0,
        top: 0,
        width: 40,
        height: 20,
      });
      shape.load("id");
      // The add's own sync: office.js hands back the id here, same as the
      // real host, though register-delay.ts still has it blocked from now.
      await context.sync();
      id = shape.id;
    });
    expect(id).not.toBe("");

    // On the current code, cleanupShapes's one-shot check sees isNullObject
    // (not gone, just not registered yet) and deletes nothing, leaving this
    // orphaned - it surfaces a moment later with nothing having claimed it.
    await cleanupShapes(slideId, [id]);

    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("a clean cleanup - every id resolves on the first check - spends no extra sync", async () => {
    const { presentation, helpers } = await bootPpt();
    await instantWait();
    const slideId = presentation.slides[0]!.id;

    let id = "";
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      const shape = shapes.addTable(1, 1, {
        left: 0,
        top: 0,
        width: 40,
        height: 20,
      });
      shape.load("id");
      await context.sync();
      id = shape.id;
    });

    const before = helpers.syncCount();
    await cleanupShapes(slideId, [id]);
    const spent = helpers.syncCount() - before;

    // The original two syncs (load+check, then delete) and nothing more: the
    // catch-up loop must never run when there is nothing left to catch up on.
    expect(spent).toBe(2);
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });
});

describe("a table whose formatting hangs never leaves the pane stuck", () => {
  it("the format sync and the cleanup queued behind it both hang: the insert still rejects", async () => {
    const { links, helpers, relay } = await bootPpt();
    await instantWait();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable([[{ t: "Q1" }, { t: "Q2" }]], [80, 80]);
    // From here: the selected slide, the slide's shapes, the add, the
    // registration poll, then the one format chunk (4); on a hung host the
    // cleanup's own first sync queues behind it and hangs too (5).
    helpers.hangNextSync(4);
    helpers.hangNextSync(5);
    vi.useFakeTimers();
    let threw: unknown;
    try {
      await settleHungSync(links.insertFromInbox(item, ws, relay));
    } catch (error) {
      threw = error;
    } finally {
      vi.useRealTimers();
    }
    expect(threw).toBeInstanceOf(Error);
    expect((threw as Error).message).toBe(
      "PowerPoint stopped answering while formatting the table",
    );
  });
});
