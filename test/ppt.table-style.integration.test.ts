// Table header and style behaviour against the fake host and the fake relay:
// the header band and PowerPoint's own default style written blind on insert,
// rebuild and repaint alike (the Mac refuses every read of a table's own
// style, so pls,fix never asks), and a repaint that clears only the fills
// pls,fix itself painted - legacy decks included. Strict load semantics are
// on, same as ppt.table.integration.test.ts beside it.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_PAINT, type TableCell } from "../src/link/model";
import { decodePaintMap, paintKey } from "../src/link/paint-map";
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

enableStrictLoadSemantics();

const WIDTHS = [80, 60];
// Row 0 is bold in every fixture below, but that never drives these tests:
// the PowerPoint side only ever acts on the `h` flag a payload already
// carries, never on the cells' own bold state (that rule is headerRow's, in
// src/excel/link-table.test.ts).
const CELLS: TableCell[][] = [
  [
    { t: "Year", b: true },
    { t: "Revenue", b: true },
  ],
  [{ t: "2024" }, { t: "1 000", f: "#EEEEEE" }],
];
// Same grid, the one fill dropped: what a repaint clears.
const REPAINTED: TableCell[][] = [
  [
    { t: "Year", b: true },
    { t: "Revenue", b: true },
  ],
  [{ t: "2025" }, { t: "1 200" }],
];
// A third row: what turns an update into a rebuild.
const GREW: TableCell[][] = [
  ...CELLS,
  [{ t: "2025" }, { t: "1 200", f: "#EEEEEE" }],
];

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

async function insertTable(cells: TableCell[][], h?: true) {
  const ws = await createWorkspace(memoryStore());
  const item = await seedTable(cells, WIDTHS, h);
  await links.insertFromInbox(item, ws, relay);
  return item;
}

function shape(index = 0): FakePptShape {
  return presentation.slides[0]!.shapes[index]!;
}

function table(index = 0): FakeTable {
  return shape(index).table!;
}

function paintTag(index = 0): Set<string> | null {
  return decodePaintMap(shape(index).tags.get(TAG_PAINT));
}

// Rolls a table back to the Mac's own no-style default, as if it had been
// inserted before table-style.ts existed: queueTableStyle overwrites every
// one of these unconditionally, so this only matters for the tests that want
// to prove a repaint fixes a table insert never touched, not for the outcome
// itself.
function forgetTableStyle(index = 0): void {
  Object.assign(table(index).styleSettings, {
    style: "",
    isFirstRowHighlighted: false,
    areRowsBanded: false,
    areColumnsBanded: false,
    isFirstColumnHighlighted: false,
    isLastRowHighlighted: false,
    isLastColumnHighlighted: false,
  });
}

describe("insert: header band and style", () => {
  it("sets the default style, the header band and no banding when the payload has a header row", async () => {
    await insertTable(CELLS, true);
    expect(table().styleSettings.style).toBe(
      PowerPoint.TableStyle.mediumStyle2Accent1,
    );
    expect(table().styleSettings.isFirstRowHighlighted).toBe(true);
    expect(table().styleSettings.areRowsBanded).toBe(false);
  });

  it("sets the default style with the header band off when the payload has none", async () => {
    await insertTable(CELLS);
    expect(table().styleSettings.style).toBe(
      PowerPoint.TableStyle.mediumStyle2Accent1,
    );
    expect(table().styleSettings.isFirstRowHighlighted).toBe(false);
  });

  it("ends at the default style no matter what PowerPoint handed the table", async () => {
    helpers.setTableStyle("shape-1", "LightStyle2Accent3");
    await insertTable(CELLS);
    expect(table().styleSettings.style).toBe(
      PowerPoint.TableStyle.mediumStyle2Accent1,
    );
  });

  it("writes the paint tag for the cells the payload filled", async () => {
    await insertTable(CELLS);
    expect(paintTag()).toEqual(new Set([paintKey(1, 1)]));
  });
});

describe("repaint: clearing only what pls,fix painted", () => {
  it("clears a fill the source dropped and leaves an always-empty cell alone", async () => {
    const item = await insertTable(CELLS);
    await pushTable(item, REPAINTED, WIDTHS);
    await links.updateLinks(await links.listLinks(relay), relay);

    expect(table().cells[1]![1]!.fill.type).toBe("NoFill");
    expect(table().cells[0]![0]!.fill.type).toBe("Inherited");
    expect(table().cells[1]![0]!.fill.type).toBe("Inherited");
  });

  it("rewrites the paint tag to the new revision's fills", async () => {
    const item = await insertTable(CELLS);
    await pushTable(item, REPAINTED, WIDTHS);
    await links.updateLinks(await links.listLinks(relay), relay);

    expect(paintTag()).toEqual(new Set());
  });
});

describe("repaint: a legacy table with no style and no paint tag", () => {
  it("sets the default style and header flag, and clears every unfilled cell except row 0", async () => {
    const item = await insertTable(CELLS, true);
    shape().tags.delete(TAG_PAINT);
    forgetTableStyle();
    await pushTable(item, REPAINTED, WIDTHS, true);
    await links.updateLinks(await links.listLinks(relay), relay);

    expect(table().styleSettings.style).toBe(
      PowerPoint.TableStyle.mediumStyle2Accent1,
    );
    expect(table().styleSettings.isFirstRowHighlighted).toBe(true);
    expect(table().cells[0]![0]!.fill.type).toBe("Inherited");
    expect(table().cells[0]![1]!.fill.type).toBe("Inherited");
    expect(table().cells[1]![0]!.fill.type).toBe("NoFill");
    expect(table().cells[1]![1]!.fill.type).toBe("NoFill");
    expect(paintTag()).toEqual(new Set());
  });
});

describe("rebuild: a table whose grid changed", () => {
  it("sets the default style and header flag, and writes the paint tag on the new shape", async () => {
    const item = await insertTable(CELLS, true);
    await pushTable(item, GREW, WIDTHS, true);
    await links.updateLinks(await links.listLinks(relay), relay);

    expect(table().rowCount).toBe(3);
    expect(table().styleSettings.style).toBe(
      PowerPoint.TableStyle.mediumStyle2Accent1,
    );
    expect(table().styleSettings.isFirstRowHighlighted).toBe(true);
    expect(paintTag()).toEqual(new Set([paintKey(1, 1), paintKey(2, 1)]));
  });
});

describe("a host below PowerPointApi 1.9", () => {
  it("never touches styleSettings, and still paints", async () => {
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.8,
    );
    await insertTable(CELLS);

    // Untouched: the Mac's own no-style default stays exactly as it is,
    // which is what queueTableStyle skips whole below TABLE_STYLE_API.
    expect(table().styleSettings.style).toBe("");
    expect(table().styleSettings.isFirstRowHighlighted).toBe(false);
    // Still paints: text, formats and the paint tag land regardless.
    expect(table().cells[0]![0]!.font.bold).toBe(true);
    expect(table().cells[1]![1]!.fill.type).toBe("Solid");
    expect(paintTag()).toEqual(new Set([paintKey(1, 1)]));
  });
});
