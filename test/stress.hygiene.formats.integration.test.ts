// Stress pass, slice P2: the older format cycles the hygiene ones sit beside
// (number, row style, fill, font colour, borders) and the two size cycles.
// These are the buttons a modeller hits hardest with a whole column selected,
// a merge in the way or a sheet somebody locked, and the two size ones run
// outside pls,fix Undo on purpose - so every abuse here is about caps, merges,
// refusals and what the undo net is allowed to claim afterwards.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import { buildSizeCycles } from "../src/cycles";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
  smt = await import("../src/excel");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

/** Every cycle in this suite's area, by the stage name it refuses under. */
function formatCycles(): [string, () => Promise<void>][] {
  return [
    ["Format cycling", () => smt.applyNumberCycle("currency")],
    ["Row styles", () => smt.applyRowStyleCycle("title")],
    ["Fill cycling", () => smt.applyFillCycle()],
    ["Font colour cycling", () => smt.applyFontColorCycle()],
    ["Border cycling", () => smt.applyBorderCycle()],
  ];
}

beforeEach(async () => {
  await boot();
});

describe("selections at and past the caps", () => {
  // 5 000 cells is the cap the grid-writing cycles count before they read
  // anything: the cell below it goes through, the cell above it is a sentence.
  it("takes the cell cap and refuses the one past it", async () => {
    helpers.select("Model!A1:A5000");
    await smt.applyNumberCycle("currency");
    expect(helpers.numberFormat("Model!A5000")).not.toBe("General");

    await boot();
    helpers.select("Model!A1:A5001");
    expect(await rejects(() => smt.applyNumberCycle("currency"))).toBe(
      "Format cycling supports up to 5,000 selected cells at once.",
    );
    expect(await rejects(() => smt.applyBorderCycle())).toBe(
      "Border cycling supports up to 5,000 selected cells at once.",
    );
  });

  it("counts the cap over every area of a ctrl-clicked selection", async () => {
    helpers.selectAreas(["Model!A1:A3000", "Model!C1:C2001"]);
    expect(await rejects(() => smt.applyNumberCycle("currency"))).toBe(
      "Format cycling supports up to 5,000 selected cells at once.",
    );
  });

  // Row styles paint per row, so their cap is rows and it is checked before
  // the undo capture: a refusal that writes nothing must not drop the stack.
  it("takes 500 rows, refuses 501 and keeps the undo stack", async () => {
    helpers.select("Model!A1:A500");
    await smt.applyRowStyleCycle("title");
    const kept = smt.undoTarget();
    expect(kept).not.toBeNull();

    helpers.select("Model!A1:A501");
    expect(await rejects(() => smt.applyRowStyleCycle("title"))).toBe(
      "Row styles support up to 500 rows at once.",
    );
    expect(smt.undoTarget()).toBe(kept);
    expect(smt.lastUndoSkipped()).toBe(false);
  });

  it("counts the row cap across areas too", async () => {
    helpers.selectAreas(["Model!A1:A250", "Model!C1:C251"]);
    expect(await rejects(() => smt.applyRowStyleCycle("title"))).toBe(
      "Row styles support up to 500 rows at once.",
    );
  });

  // A whole-sheet Ctrl+A is cellCount -1, which reads as the largest selection
  // there is rather than the smallest.
  it("refuses Ctrl+A on every cycle that reads or writes a grid", async () => {
    for (const [stage, run] of [
      ["Format cycling", () => smt.applyNumberCycle("currency")],
      ["Border cycling", () => smt.applyBorderCycle()],
    ] as [string, () => Promise<void>][]) {
      await boot();
      helpers.select("Model!A1:XFD1048576");
      expect(await rejects(run)).toBe(
        `${stage} supports up to 5,000 selected cells at once.`,
      );
    }

    await boot();
    helpers.select("Model!A1:XFD1048576");
    expect(await rejects(() => smt.applyRowStyleCycle("title"))).toBe(
      "Row styles support up to 500 rows at once.",
    );
  });
});

describe("merged cells in the way", () => {
  it("formats a whole merged block", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:C1");
    helpers.select("Model!A1:C1");

    await smt.applyNumberCycle("currency");
    await smt.applyFillCycle();
    await smt.applyBorderCycle();
    expect(helpers.value("Model!A1")).toBe("Heading");
    expect(helpers.sheet("Model").merges).toHaveLength(1);
  });

  // Excel refuses a VALUE write that cuts a merge, not a format one, so these
  // cycles go through on a half-in block. What matters is that the merge is
  // still standing and the heading still reads.
  it("formats a block that cuts a merge without breaking it", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:C1");
    helpers.select("Model!B1:D1");

    await smt.applyNumberCycle("currency");
    await smt.applyFillCycle();
    await smt.applyBorderCycle();
    expect(helpers.sheet("Model").merges).toHaveLength(1);
    expect(helpers.value("Model!A1")).toBe("Heading");
  });
});

describe("cells a cycle was not written for", () => {
  it("cycles the number format over text, errors, booleans and blanks", async () => {
    helpers.seed("Model!A1", [
      ["text"],
      [{ formula: "=NA()", value: "#N/A" }],
      [{ formula: "=A99", value: "#REF!" }],
      [true],
      [0],
      [-1234.5],
      [""],
    ]);
    helpers.select("Model!A1:A7");

    await smt.applyNumberCycle("currency");
    expect(helpers.value("Model!A1")).toBe("text");
    expect(helpers.value("Model!A4")).toBe(true);
    expect(helpers.formula("Model!A3")).toBe("=A99");
    expect(helpers.numberFormat("Model!A7")).toBe(
      helpers.numberFormat("Model!A1"),
    );
  });

  it("walks the date ladder round and back to where it started", async () => {
    helpers.seed("Model!A1", [[45000]]);
    helpers.select("Model!A1");

    const first = helpers.numberFormat("Model!A1");
    const walk: string[] = [];
    for (let press = 0; press < 4; press += 1) {
      await smt.applyNumberCycle("date");
      walk.push(helpers.numberFormat("Model!A1"));
    }
    expect(new Set(walk).size).toBe(3);
    expect(walk[3]).toBe(walk[0]);
    expect(first).not.toBe(walk[0]);
  });

  it("steps the fill cycle off a hand-painted colour and back to clear", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    helpers.setFill("Model!A1", { color: "#123456", pattern: "Solid" });

    const walk: string[] = [];
    for (let press = 0; press < 5; press += 1) {
      await smt.applyFillCycle();
      walk.push(helpers.fill("Model!A1").pattern);
    }
    expect(walk).toContain("None");
    expect(helpers.value("Model!A1")).toBe("Revenue");
  });
});

describe("a host that says no", () => {
  it("answers a protected sheet with the stage's own sentence", async () => {
    for (const [stage, run] of formatCycles()) {
      await boot();
      helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
      helpers.select("Model!A1:A2");
      helpers.protectSheet("Model");
      const before = helpers.cellMap("Model");

      expect(await rejects(run)).toBe(
        `${stage}: this sheet is protected, nothing was changed`,
      );
      expect(helpers.cellMap("Model")).toEqual(before);
    }
  });

  it("names the build when a ctrl-clicked selection meets a host below 1.9", async () => {
    for (const [stage, run] of formatCycles()) {
      await boot({
        isSetSupported: (set, version) =>
          !(set === "ExcelApi" && version === "1.9"),
      });
      helpers.seed("Model!A1", [["Revenue"]]);
      helpers.selectAreas(["Model!A1", "Model!C1"]);

      expect(await rejects(run)).toBe(
        `${stage}: this Excel build can only act on one selected block.`,
      );
    }
  });

  it("leaves the cells alone when the host refuses the batch", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    const before = helpers.cellMap("Model");
    helpers.failNextSync();

    await rejects(() => smt.applyFillCycle());
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});

describe("the two size cycles", () => {
  it("walks the row-height and column-width ladders and comes home", async () => {
    const ladders = buildSizeCycles();
    helpers.select("Model!A1:B2");

    const heights: number[] = [];
    for (let press = 0; press < ladders.rowHeight.length + 1; press += 1) {
      await smt.applyRowHeightCycle();
      heights.push(helpers.rowHeight("Model", 0));
    }
    expect(heights.slice(0, ladders.rowHeight.length)).toEqual([
      ...ladders.rowHeight.slice(1),
      ladders.rowHeight[0],
    ]);

    const widths: number[] = [];
    for (let press = 0; press < ladders.columnWidth.length; press += 1) {
      await smt.applyColumnWidthCycle();
      widths.push(helpers.columnWidth("Model", 0));
    }
    expect(new Set(widths).size).toBe(ladders.columnWidth.length);
  });

  // A hand-dragged size is on no rung of ours, so the next press starts the
  // ladder rather than guessing which rung the modeller meant.
  it("starts the ladder from a hand-dragged size", async () => {
    const ladders = buildSizeCycles();
    helpers.select("Model!A1");
    await smt.applyRowHeightCycle();
    await smt.applyRowHeightCycle();
    expect(helpers.rowHeight("Model", 0)).toBe(ladders.rowHeight[2]);
  });

  it("refuses a ctrl-clicked selection by name", async () => {
    helpers.selectAreas(["Model!A1", "Model!C1"]);
    expect(await rejects(() => smt.applyRowHeightCycle())).toBe(
      "Row height: select a single range",
    );
    expect(await rejects(() => smt.applyColumnWidthCycle())).toBe(
      "Column width: select a single range",
    );
  });

  it("answers a protected sheet without touching a cell", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1:B2");
    helpers.protectSheet("Model");
    const before = helpers.cellMap("Model");

    expect(await rejects(() => smt.applyRowHeightCycle())).toBe(
      "Row height: this sheet is protected, nothing was changed",
    );
    expect(await rejects(() => smt.applyColumnWidthCycle())).toBe(
      "Column width: this sheet is protected, nothing was changed",
    );
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  // Sizes are sheet state; getCellProperties carries none of it, so these two
  // never capture and never claim the net is off either.
  it("spends no undo slot and never reports one skipped", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    await smt.applyRowHeightCycle();
    await smt.applyColumnWidthCycle();

    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(false);
  });

  it("sizes a merged block without breaking the merge", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:C1");
    helpers.select("Model!A1:C1");

    await smt.applyRowHeightCycle();
    await smt.applyColumnWidthCycle();
    expect(helpers.sheet("Model").merges).toHaveLength(1);
    expect(helpers.value("Model!A1")).toBe("Heading");
  });

  it("still sizes a single block on a host below ExcelApi 1.9", async () => {
    await boot({
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.9"),
    });
    helpers.select("Model!A1:B2");
    await smt.applyRowHeightCycle();
    expect(helpers.rowHeight("Model", 0)).not.toBe(15);
  });
});
