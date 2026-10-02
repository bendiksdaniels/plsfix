// Hunt pass 2, target 4: overTableCap (link/model.ts) checks only the
// ceiling. requireTableSize (src/ppt/tables.ts:63) forwarded straight to it,
// so a hand-crafted or corrupted TablePayload with 0 rows or 0 columns sailed
// through both insertTable and refreshTable's guard and reached
// shapes.addTable, which throws a raw InvalidArgument - a raw Office code in
// the toast, and (on a repaint) a table left in whatever state the host's
// own refusal happened to leave it in. Fixed at the root in tables.ts:
// requireTableSize now refuses first, before any PowerPoint.run.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { TableCell, TablePayload } from "../../src/link/model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "../fakeppt";
import { bootPpt, memoryStore, pushTable, seedTable } from "../ppt.support";
import type * as LinksModule from "../../src/ppt/links";
import { requireTableSize } from "../../src/ppt/tables";

enableStrictLoadSemantics();

const GOOD_CELLS: TableCell[][] = [
  [{ t: "Q1" }, { t: "Q2" }],
  [{ t: "100" }, { t: "120" }],
];
const GOOD_WIDTHS = [80, 80];

function payloadWith(rows: number, cols: number): TablePayload {
  return {
    v: 1,
    kind: "table",
    rows,
    cols,
    cells: GOOD_CELLS,
    widths: GOOD_WIDTHS,
    src: { workbook: "Model_v4.xlsx", sheet: "Model", ref: "A1", anchor: "x" },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

// No raw Office code ever reaches the toast: the same sentinel every other
// hunt file in this folder checks a plain-sentence error against.
function expectPlainSentence(line: string | undefined): void {
  expect(line).toBeDefined();
  expect(line).not.toMatch(/InvalidArgument|ItemNotFound|GeneralException/);
}

describe("requireTableSize: the lower bound overTableCap never checked", () => {
  it.each([
    ["0 rows", 0, 2],
    ["0 columns", 2, 0],
    ["0 rows and 0 columns", 0, 0],
    ["a negative row count", -1, 2],
    ["a non-integer row count", 1.5, 2],
  ] as const)(
    "refuses %s with a plain sentence, not TABLE_TOO_BIG's",
    (_label, rows, cols) => {
      expect(() => requireTableSize(payloadWith(rows, cols))).toThrow(
        "Tables need at least one row and one column.",
      );
    },
  );

  it("still refuses past the ceiling with the existing sentence", () => {
    expect(() => requireTableSize(payloadWith(61, 21))).toThrow(
      "Tables go up to 60 rows and 20 columns; export a picture for more.",
    );
  });

  it("accepts the smallest real table, 1x1", () => {
    expect(() => requireTableSize(payloadWith(1, 1))).not.toThrow();
  });
});

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
  uninstallFakePpt();
});

describe("inserting a table export with 0 rows or 0 columns", () => {
  it.each([
    ["0 rows", [] as TableCell[][], GOOD_WIDTHS],
    // A structurally CONSISTENT 0-column grid (two rows, each of length
    // zero, matching widths.length): isTablePayload's own isCellGrid check
    // only compares shapes against each other, never against a minimum, so
    // this reaches tables.ts as a "valid" payload - unlike mismatched cells
    // against a shorter widths array, which decodePayload already refuses
    // as corrupt before this guard ever runs.
    ["0 columns", [[], []] as TableCell[][], [] as number[]],
  ] as const)(
    "%s refuses before any PowerPoint.run, no shape left behind",
    async (_label, cells, widths) => {
      const item = await seedTable(cells, widths);
      const before = helpers.syncCount();

      await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
        "Tables need at least one row and one column.",
      );

      // The guard fires before resolveTarget or addTable ever run: zero syncs
      // spent, not merely zero shapes landed.
      expect(helpers.syncCount()).toBe(before);
      expect(presentation.slides[0]!.shapes).toHaveLength(0);
    },
  );
});

describe("the fake host models real PowerPoint's own refusal, for the case the guard is bypassed", () => {
  it("shapes.addTable(0, cols) throws InvalidArgument, never a silent empty grid", async () => {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(
        presentation.slides[0]!.id,
      ).shapes;
      expect(() =>
        shapes.addTable(0, 2, { left: 0, top: 0, width: 100, height: 50 }),
      ).toThrow(/InvalidArgument/);
      expect(() =>
        shapes.addTable(2, 0, { left: 0, top: 0, width: 100, height: 50 }),
      ).toThrow(/InvalidArgument/);
      await context.sync();
    });
  });
});

describe("a relay push corrupted to 0 rows, fetched by Update all", () => {
  it("fails the row with a plain sentence and leaves the old table untouched", async () => {
    const item = await seedTable(GOOD_CELLS, GOOD_WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    const before = (await links.listLinks(relay))[0]!;
    const beforeTable = presentation.slides[0]!.shapes[0]!.table;
    expect(beforeTable?.rowCount).toBe(2);

    // A corrupted or hand-edited push: the source's own export path never
    // produces this (Excel always ships rows === cells.length), but a relay
    // row from another build or a hand-crafted payload can.
    await pushTable(item, [], GOOD_WIDTHS);
    const rows = await links.listLinks(relay);
    expect(rows).toHaveLength(1);

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expectPlainSentence(summary.failures[0]);
    expect(summary.failures[0]).toContain(
      "Tables need at least one row and one column.",
    );
    // The shape the deck shows is still the OLD table: refreshTable's own
    // requireTableSize call refuses before touching rowCount, cells or tags.
    const afterTable = presentation.slides[0]!.shapes[0]!.table;
    expect(afterTable?.rowCount).toBe(2);
    expect(afterTable?.cells[0]?.[0]?.text).toBe("Q1");
    expect((await links.listLinks(relay))[0]!.found.tag.rev).toBe(
      before.found.tag.rev,
    );
  });
});
