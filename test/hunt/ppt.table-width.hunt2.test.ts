// Hunt pass 2: PowerPoint for the web refuses ShapeCollection.addTable with
// InvalidArgument whenever the columns' columnWidth values do not add up
// EXACTLY to the table's own width (probed 27.09; the fake's own check is
// test/fakeppt/size-guard.ts requireColumnWidthsMatch). fitInto (src/layout.ts)
// never scales up, so a table whose source columns sum to a fractional width
// lands with a fractional box width whenever it is narrower than its spot -
// on insert through a named Spot (src/ppt/placement.ts's spotTarget, which
// calls fitInto directly, unlike the free-space search's own rounding) or on
// a rebuild at an existing shape's own (possibly fractional) width.

import { afterEach, describe, expect, it } from "vitest";
import { createWorkspace } from "../../src/link/workspace";
import type { TableCell } from "../../src/link/model";
import type { InsertTarget } from "../../src/ppt/placement";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
} from "../fakeppt";
import { bootPpt, memoryStore, pushTable, seedTable } from "../ppt.support";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

function columnWidthSum(shape: FakePptShape): number {
  return shape.table!.columnWidths.reduce<number>(
    (total, one) => total + (one ?? 0),
    0,
  );
}

describe("insert: a table link whose source columns sum to a fractional width", () => {
  // 60.75 + 48 + 48 + 48 = 204.75: the exact shape of a real Excel export
  // (one wide label column, three even data columns) that only comes out
  // fractional because of the label column's own odd measurement.
  const FRACTIONAL_WIDTHS = [60.75, 48, 48, 48];
  const CELLS: TableCell[][] = [
    [{ t: "Label" }, { t: "2024" }, { t: "2025" }, { t: "2026" }],
  ];

  it("lands inside a chosen spot with its width equal to its columns' sum, not InvalidArgument", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, FRACTIONAL_WIDTHS);
    // A named spot goes through placement.ts's spotTarget, which calls
    // layout.ts's fitInto directly - unlike "free", the default, whose own
    // free-space search (src/free-space.ts) rounds every box it hands back.
    // The spot is roomy enough that the table lands at its own fractional
    // size, unscaled: fitInto never enlarges, so scale is 1 here and the
    // box's width is the source columns' own fractional total, unrounded.
    const target: InsertTarget = {
      slideId: presentation.slides[0]!.id,
      where: "top-right",
    };

    const placed = await links.insertFromInbox(item, ws, relay, target);

    const shape = presentation.slides[0]!.shapes.find(
      (one) => one.id === placed.shapeId,
    )!;
    expect(shape.type).toBe("Table");
    expect(shape.width).toBe(Math.round(204.75));
    expect(columnWidthSum(shape)).toBe(shape.width);
  });
});

describe("rebuild: a table whose live width in the deck is already fractional", () => {
  const WIDTHS = [80, 60];
  const CELLS: TableCell[][] = [[{ t: "Revenue" }, { t: "1 000" }]];
  const GREW: TableCell[][] = [...CELLS, [{ t: "Costs" }, { t: "-400" }]];

  it("rebuilds at that width without InvalidArgument, its columns summing to the rounded width", async () => {
    const { links, presentation, relay } = await bootPpt();
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const before = presentation.slides[0]!.shapes[0]!;
    // A live PowerPoint width is not always a whole number - dragged by hand,
    // or left over from a table built before this fix - and recreate() reads
    // it as-is, with no fitInto and no free-space rounding in between.
    before.width = 149.6;
    // A row added: rowCount no longer matches, so refreshTable takes the
    // rebuild branch (recreate), at this same corner and width.
    await pushTable(item, GREW, WIDTHS);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const after = presentation.slides[0]!.shapes[0]!;
    expect(after.type).toBe("Table");
    expect(after.width).toBe(Math.round(149.6));
    expect(columnWidthSum(after)).toBe(after.width);
  });
});
