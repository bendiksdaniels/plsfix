// Native table links against the fake host and the fake relay: an insert
// builds a real Table shape in the slide's free space with its cell formats
// and both tags, an update rewrites the cells where the table sits, a source
// that changed shape has the table built again at the same corner, and a host
// below PowerPointApi 1.8 says so instead of inserting. Strict load semantics
// are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { overlaps } from "../src/layout";
import { TAG_KEY, TAG_LINK, type TableCell } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
  type FakeTable,
} from "./fakeppt";
import { bootPpt, memoryStore, pushTable, seedTable } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type * as RevertModule from "../src/ppt/revert";

enableStrictLoadSemantics();

const WIDTHS = [80, 60];
const CELLS: TableCell[][] = [
  [
    { t: "Revenue", b: true },
    { t: "1 000", a: "r", z: 9 },
  ],
  [
    { t: "Costs", i: true, c: "#FF0000" },
    { t: "-400", f: "#EEEEEE" },
  ],
];
const AFTER: TableCell[][] = [
  [
    { t: "Revenue", b: true },
    { t: "1 200", a: "r", z: 9 },
  ],
  [
    { t: "Costs", i: true, c: "#FF0000" },
    { t: "-450", f: "#EEEEEE" },
  ],
];
const GREW: TableCell[][] = [...AFTER, [{ t: "Profit" }, { t: "750" }]];

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

async function insertTable(
  cells: TableCell[][] = CELLS,
): Promise<Awaited<ReturnType<typeof seedTable>>> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedTable(cells, WIDTHS);
  await links.insertFromInbox(item, ws, relay);
  return item;
}

function shape(index = 0): FakePptShape {
  return presentation.slides[0]!.shapes[index]!;
}

function table(index = 0): FakeTable {
  return shape(index).table!;
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

describe("insert a table link", () => {
  it("builds a Table shape in the free space, with the cell formats and both tags", async () => {
    const item = await insertTable();
    const inserted = shape();
    expect(inserted.type).toBe("Table");
    expect(inserted.name).toBe(`pls,fix table ${item.label}`);
    // The columns' own widths, and a row's worth of height per row, at the
    // first free spot on an empty slide.
    expect(box(inserted)).toEqual({
      left: 410,
      top: 252,
      width: 140,
      height: 36,
    });

    const grid = table();
    expect(grid.rowCount).toBe(2);
    expect(grid.columnCount).toBe(2);
    expect(grid.cells.map((row) => row.map((cell) => cell.text))).toEqual([
      ["Revenue", "1 000"],
      ["Costs", "-400"],
    ]);
    expect(grid.cells[0]![0]!.font.bold).toBe(true);
    expect(grid.cells[0]![1]!.horizontalAlignment).toBe("Right");
    expect(grid.cells[0]![1]!.font.size).toBe(9);
    expect(grid.cells[1]![0]!.font.italic).toBe(true);
    expect(grid.cells[1]![0]!.font.color).toBe("#FF0000");
    expect(grid.cells[1]![1]!.fill.color).toBe("#EEEEEE");
    // A cell with nothing to format is left to the deck's table style.
    expect(grid.cells[1]![1]!.font.bold).toBeNull();

    expect(JSON.parse(inserted.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "table",
      rev: 1,
    });
    expect(inserted.tags.get(TAG_KEY)).toBe(item.token);
    expect(
      await links.listInbox(await createWorkspace(memoryStore()), relay),
    ).toHaveLength(0);
  });

  it("lands a second table beside the first, not on it", async () => {
    await insertTable();
    await insertTable();
    expect(overlaps(box(shape(0)), box(shape(1)))).toBe(false);
  });

  it("says which PowerPoint a table needs, and inserts nothing", async () => {
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
    await expect(insertTable()).rejects.toThrow(
      "Tables need PowerPoint 2021 or Microsoft 365.",
    );
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });
});

// Rows come out double height on insert, and only shrink once a repaint's own
// font.size writes land, when the table is created at PowerPoint's default
// (18 pt) and resized after the fact. Creating it with the payload's own
// modal size already set avoids that first-paint overflow.
describe("insert: uniform font size", () => {
  it("creates the table with the modal cell size, and still writes a cell whose own size differs", async () => {
    const cells: TableCell[][] = [
      [
        { t: "Year", z: 11 },
        { t: "Revenue", z: 11 },
      ],
      [
        { t: "2024", z: 11 },
        { t: "1 000", z: 9 },
      ],
    ];
    await insertTable(cells);
    expect(table().uniformCellProperties).toEqual({ font: { size: 11 } });
    expect(table().cells[0]![0]!.font.size).toBe(11);
    expect(table().cells[0]![1]!.font.size).toBe(11);
    expect(table().cells[1]![0]!.font.size).toBe(11);
    // The minority cell still gets its own size: it is not left at the
    // uniform one the table was created with.
    expect(table().cells[1]![1]!.font.size).toBe(9);
  });

  it("names no uniform size when the payload carries none", async () => {
    const cells: TableCell[][] = [
      [{ t: "Year" }, { t: "Revenue" }],
      [{ t: "2024" }, { t: "1 000" }],
    ];
    await insertTable(cells);
    expect(table().uniformCellProperties).toBeNull();
  });
});

describe("update a table link", () => {
  it("rewrites the cells where the table sits and leaves its geometry alone", async () => {
    const item = await insertTable();
    const moved = shape();
    moved.left = 200;
    moved.top = 90;
    moved.width = 300;
    moved.height = 80;
    await pushTable(item, AFTER, WIDTHS);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(shape().id).toBe(moved.id);
    expect(box(shape())).toEqual({
      left: 200,
      top: 90,
      width: 300,
      height: 80,
    });
    expect(table().cells.map((row) => row.map((cell) => cell.text))).toEqual([
      ["Revenue", "1 200"],
      ["Costs", "-450"],
    ]);
    expect(JSON.parse(shape().tags.get(TAG_LINK)!).rev).toBe(2);
  });

  // A cell that was bold and is not any more has to lose the bold: a repaint
  // writes the whole format, not only what the payload names.
  it("clears a format the source dropped", async () => {
    const item = await insertTable();
    await pushTable(
      item,
      [[{ t: "Revenue" }, { t: "1 000" }], AFTER[1]!],
      WIDTHS,
    );
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(table().cells[0]![0]!.font.bold).toBe(false);
    expect(table().cells[0]![1]!.horizontalAlignment).toBe("Left");
    expect(table().cells[1]![1]!.fill.color).toBe("#EEEEEE");
  });

  it("builds the table again at the same corner when the source grew, and re-tags it", async () => {
    const item = await insertTable();
    const before = shape();
    before.left = 120;
    before.top = 60;
    before.width = 260;
    await pushTable(item, GREW, WIDTHS);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    const after = shape();
    expect(after.id).not.toBe(before.id);
    expect(after.type).toBe("Table");
    expect([after.left, after.top, after.width]).toEqual([120, 60, 260]);
    expect(after.height).toBe(3 * 18);
    expect(table().rowCount).toBe(3);
    expect(table().cells[2]!.map((cell) => cell.text)).toEqual([
      "Profit",
      "750",
    ]);
    expect(JSON.parse(after.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "table",
      rev: 2,
    });
    expect(after.tags.get(TAG_KEY)).toBe(item.token);
    expect(await links.listLinks(relay)).toHaveLength(1);
  });

  it("puts the older cells back on a revert", async () => {
    const revert: typeof RevertModule = await import("../src/ppt/revert");
    const item = await insertTable();
    await pushTable(item, AFTER, WIDTHS);
    await links.updateLinks(await links.listLinks(relay), relay);

    const summary = await revert.revertLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ reverted: 1, failed: 0 });
    expect(table().cells.map((row) => row.map((cell) => cell.text))).toEqual([
      ["Revenue", "1 000"],
      ["Costs", "-400"],
    ]);
    expect(JSON.parse(shape().tags.get(TAG_LINK)!).rev).toBe(1);
  });
});
