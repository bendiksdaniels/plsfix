// Table cells are written CELLS_PER_SYNC per round trip: a 6x4 repaint is a
// read and three chunks, the tag lands with the last one, a chunk the host
// swallows leaves the old tag standing, and an insert whose format round trip
// stops answering takes its half-formatted table down again.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_LINK, type TableCell } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import { settleHungSync } from "./hung-sync";
import { CELLS_PER_SYNC } from "../src/ppt/tables";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushTable, seedTable } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const WIDTHS = [80, 60, 60, 60];

function grid(prefix: string, rows = 6, cols = 4): TableCell[][] {
  return Array.from({ length: rows }, (_row, r) =>
    Array.from({ length: cols }, (_col, c) => ({ t: `${prefix}${r}${c}` })),
  );
}

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

function tables(slide = 0): FakePptShape[] {
  return presentation.slides[slide]!.shapes.filter((s) => s.type === "Table");
}

function cellText(shape: FakePptShape, row: number, col: number): string {
  return shape.table!.cells[row]![col]!.text;
}

describe("a table repaint", () => {
  it("is a read and one round trip per chunk, the tag with the last", async () => {
    expect(CELLS_PER_SYNC).toBe(8);
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(grid("a"), WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const shape = tables()[0]!;
    const tagBefore = shape.tags.get(TAG_LINK);
    await pushTable(item, grid("b"), WIDTHS);
    const rows = await links.listLinks(relay);

    const before = helpers.syncCount();
    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    // 24 cells at 8 per round trip: three chunks after the dimension read.
    expect(helpers.syncCount() - before).toBe(1 + 3);
    expect(cellText(shape, 0, 0)).toBe("b00");
    expect(cellText(shape, 5, 3)).toBe("b53");
    expect(shape.tags.get(TAG_LINK)).not.toBe(tagBefore);
  });

  it("keeps the old tag when the host stops answering mid-way", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(grid("a"), WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const shape = tables()[0]!;
    const tagBefore = shape.tags.get(TAG_LINK);
    await pushTable(item, grid("b"), WIDTHS);
    const rows = await links.listLinks(relay);

    // The read is the first round trip, the second chunk the third.
    helpers.hangNextSync(2);
    vi.useFakeTimers();
    const summary = await settleHungSync(links.updateLinks(rows, relay));

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(
      /stopped answering while repainting the table/,
    );
    // The first chunk landed, the tag did not: the row stays updatable and
    // the next update writes every cell again.
    expect(cellText(shape, 0, 0)).toBe("b00");
    expect(shape.tags.get(TAG_LINK)).toBe(tagBefore);

    vi.useRealTimers();
    const again = await links.updateLinks(await links.listLinks(relay), relay);
    expect(again).toMatchObject({ updated: 1, failed: 0 });
    expect(cellText(shape, 5, 3)).toBe("b53");
    expect(shape.tags.get(TAG_LINK)).not.toBe(tagBefore);
  });
});

describe("a table insert whose format round trip stops answering", () => {
  it("leaves no table behind, and the next insert lands", async () => {
    const ws = await createWorkspace(memoryStore());
    const bold = grid("a").map((row) =>
      row.map((cell) => ({ ...cell, b: true as const })),
    );
    const item = await seedTable(bold, WIDTHS);
    const before = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    const perInsert = helpers.syncCount() - before;
    // 24 formatted cells: the insert round trip plus three format chunks.
    expect(perInsert).toBeGreaterThanOrEqual(1 + 3);
    expect(tables()).toHaveLength(1);

    const second = await seedTable(bold, WIDTHS);
    helpers.selectSlide(presentation.slides[1]!.id);
    // The last round trip of the same insert is the last format chunk.
    helpers.hangNextSync(perInsert - 1);
    vi.useFakeTimers();
    await expect(
      settleHungSync(links.insertFromInbox(second, ws, relay)),
    ).rejects.toThrow(/stopped answering while formatting the table/);
    vi.useRealTimers();

    expect(tables(1)).toHaveLength(0);
    const third = await seedTable(bold, WIDTHS);
    await links.insertFromInbox(third, ws, relay);
    expect(tables(1)).toHaveLength(1);
  });
});
