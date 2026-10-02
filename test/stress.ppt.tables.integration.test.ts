// Stress pass on table links: a source that grows to the 60 x 20 cap and past
// it, and a rebuild whose format round trip the host swallows. Invariant: a
// table is never left half formatted while the pane calls it up to date, and a
// payload past the cap is refused in words instead of spending a round trip
// per eight cells.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  TABLE_MAX_COLS,
  TABLE_MAX_ROWS,
  TABLE_TOO_BIG,
  TAG_LINK,
  type TableCell,
} from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { Workspace } from "../src/link/workspace";
import { settleHungSync } from "./hung-sync";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushTable, seedTable } from "./ppt.support";
import { CELLS_PER_SYNC } from "../src/ppt/tables";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

// A grid whose every cell carries a format, so each one costs a write and the
// chunking is what a repaint actually spends.
function bold(prefix: string, rows: number, cols: number): TableCell[][] {
  return Array.from({ length: rows }, (_row, r) =>
    Array.from({ length: cols }, (_col, c) => ({
      t: `${prefix}${String(r)}-${String(c)}`,
      b: true as const,
    })),
  );
}

function widths(cols: number): number[] {
  return Array.from({ length: cols }, () => 40);
}

function tables(slide = 0): FakePptShape[] {
  return presentation.slides[slide]!.shapes.filter(
    (shape) => shape.type === "Table",
  );
}

function rev(shape: FakePptShape): number {
  return (JSON.parse(shape.tags.get(TAG_LINK)!) as { rev: number }).rev;
}

async function failureOf(run: Promise<unknown>): Promise<string> {
  try {
    await run;
    throw new Error("expected a refusal");
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}

describe("a table source that grows to the cap", () => {
  it("rebuilds a 6x4 into the full 60 x 20 and writes every cell", async () => {
    const item = await seedTable(bold("a", 6, 4), widths(4));
    await links.insertFromInbox(item, ws, relay);
    await pushTable(
      item,
      bold("b", TABLE_MAX_ROWS, TABLE_MAX_COLS),
      widths(TABLE_MAX_COLS),
    );

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(tables()).toHaveLength(1);
    const grid = tables()[0]!.table!;
    expect([grid.rowCount, grid.columnCount]).toEqual([
      TABLE_MAX_ROWS,
      TABLE_MAX_COLS,
    ]);
    expect(grid.cells[59]![19]!.text).toBe("b59-19");
    expect(grid.cells[59]![19]!.font.bold).toBe(true);
    expect(await links.listLinks(relay)).toMatchObject([{ status: "current" }]);
  });

  // A characterization test, not a rule: 60 x 20 is inside the documented cap
  // and costs 151 round trips, which on the web's bad day (0.4 s per property
  // write, lessons 09.09) is minutes. Any change to CELLS_PER_SYNC or to the
  // cap has to move this number on purpose.
  it("spends one round trip per eight cells repainting at the cap", async () => {
    const plain = (prefix: string): TableCell[][] =>
      Array.from({ length: TABLE_MAX_ROWS }, (_row, r) =>
        Array.from({ length: TABLE_MAX_COLS }, (_col, c) => ({
          t: `${prefix}${String(r)}-${String(c)}`,
        })),
      );
    const item = await seedTable(plain("a"), widths(TABLE_MAX_COLS));
    await links.insertFromInbox(item, ws, relay);
    await pushTable(item, plain("b"), widths(TABLE_MAX_COLS));
    const rows = await links.listLinks(relay);

    const before = helpers.syncCount();
    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const chunks = Math.ceil(
      (TABLE_MAX_ROWS * TABLE_MAX_COLS) / CELLS_PER_SYNC,
    );
    expect(chunks).toBe(150);
    // The dimension read, then the cells.
    expect(helpers.syncCount() - before).toBe(1 + chunks);
  });

  it("refuses a source one row and one column past the cap, both ways", async () => {
    const over = await seedTable(
      bold("x", TABLE_MAX_ROWS + 1, TABLE_MAX_COLS + 1),
      widths(TABLE_MAX_COLS + 1),
    );
    const from = helpers.syncCount();

    expect(await failureOf(links.insertFromInbox(over, ws, relay))).toContain(
      TABLE_TOO_BIG,
    );

    // Refused before anything is drawn, so the slide is untouched and the
    // 1 281 cells never cost a round trip each.
    expect(tables()).toHaveLength(0);
    expect(helpers.syncCount() - from).toBe(0);

    // And a link already in the deck whose source outgrew the cap says the
    // same, rather than repainting for minutes.
    const item = await seedTable(bold("a", 2, 2), widths(2));
    await links.insertFromInbox(item, ws, relay);
    await pushTable(
      item,
      bold("b", TABLE_MAX_ROWS + 1, TABLE_MAX_COLS),
      widths(TABLE_MAX_COLS),
    );
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain(TABLE_TOO_BIG);
    expect(tables()[0]!.table!.rowCount).toBe(2);
  });
});

describe("a rebuild the host stops answering", () => {
  it("keeps the row updatable, and the next update finishes the formats", async () => {
    const item = await seedTable(bold("a", 6, 4), widths(4));
    await links.insertFromInbox(item, ws, relay);
    const before = rev(tables()[0]!);
    await pushTable(item, bold("b", 7, 4), widths(4));
    const rows = await links.listLinks(relay);

    // The read is the first round trip, the rebuild the second, and the
    // formats follow eight cells at a time: this is the second of those.
    helpers.hangNextSync(3);
    vi.useFakeTimers();
    const summary = await settleHungSync(links.updateLinks(rows, relay));
    vi.useRealTimers();

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(
      /stopped answering while formatting the table/,
    );
    // The new grid is there but only half of it is formatted, so the tag must
    // still hold the old revision: a table nobody can finish formatting is
    // worse than one the next press repaints.
    const half = tables()[0]!.table!;
    expect(half.rowCount).toBe(7);
    expect(half.cells[0]![0]!.font.bold).toBe(true);
    expect(half.cells[6]![3]!.font.bold).toBeNull();
    expect(rev(tables()[0]!)).toBe(before);
    expect(await links.listLinks(relay)).toMatchObject([
      { status: "updateAvailable" },
    ]);

    const again = await links.updateLinks(await links.listLinks(relay), relay);
    expect(again).toMatchObject({ updated: 1, failed: 0 });
    const healed = tables()[0]!.table!;
    expect(healed.cells[6]![3]!.font.bold).toBe(true);
    expect(healed.cells[6]![3]!.text).toBe("b6-3");
    expect(rev(tables()[0]!)).toBe(before + 1);
  });

  it("leaves one table on the slide, never the old one beside the new", async () => {
    const item = await seedTable(bold("a", 6, 4), widths(4));
    await links.insertFromInbox(item, ws, relay);
    await pushTable(item, bold("b", 7, 4), widths(4));
    const rows = await links.listLinks(relay);

    helpers.hangNextSync(3);
    vi.useFakeTimers();
    await settleHungSync(links.updateLinks(rows, relay));
    vi.useRealTimers();

    expect(tables()).toHaveLength(1);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });
});

describe("a table the user rearranged", () => {
  it("says to ungroup it before its size can change", async () => {
    const item = await seedTable(bold("a", 2, 2), widths(2));
    await links.insertFromInbox(item, ws, relay);
    const table = tables()[0]!;
    const caption = presentation.addShape(presentation.slides[0]!, {
      left: 5,
      top: 5,
      width: 60,
    });
    presentation.groupShapes(
      [table.id, caption.id],
      presentation.slides[0]!.id,
    );
    await pushTable(item, bold("b", 3, 2), widths(2));

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain(
      "ungroup the table before its size can change",
    );
    // Nothing was rebuilt: the grouped table still holds what it held.
    expect(table.table!.rowCount).toBe(2);
  });

  it("repaints a grouped table whose grid did not change", async () => {
    const item = await seedTable(bold("a", 2, 2), widths(2));
    await links.insertFromInbox(item, ws, relay);
    const table = tables()[0]!;
    const caption = presentation.addShape(presentation.slides[0]!, {
      left: 5,
      top: 5,
      width: 60,
    });
    presentation.groupShapes(
      [table.id, caption.id],
      presentation.slides[0]!.id,
    );
    await pushTable(item, bold("b", 2, 2), widths(2));

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(table.table!.cells[1]![1]!.text).toBe("b1-1");
  });
});
