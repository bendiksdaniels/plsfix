// Stress pass, slice P2: the three hygiene cycles (indent, alignment,
// underline) pressed with the wrong thing selected, at the wrong time, twice.
// Each one reads one scalar off the active cell, steps a ladder in
// src/cycles.ts and writes into every area, so the abuse that matters is a
// read-back Excel rewrote, a state no rung owns, and a write the host refuses.
// Every path must end on a pane sentence and a workbook that is whole.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

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

/** The three cycles under one name, so a scenario runs against all of them. */
function hygieneCycles(): [string, () => Promise<string>][] {
  return [
    ["Indent cycling", () => smt.applyIndentCycle()],
    ["Alignment cycling", () => smt.applyAlignmentCycle()],
    ["Underline cycling", () => smt.applyUnderlineCycle()],
  ];
}

beforeEach(async () => {
  await boot();
});

describe("a read-back no rung owns", () => {
  // Excel's maximum indent in the UI. It is not on our ladder, so the press
  // has to start the ladder over rather than throw or sit still.
  it("steps home from Excel's maximum indent of 15", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    helpers.setIndent("Model!A1", 15);

    await smt.applyIndentCycle();
    expect(helpers.cell("Model!A1").indentLevel).toBe(0);
    await smt.applyIndentCycle();
    expect(helpers.cell("Model!A1").indentLevel).toBe(1);
  });

  it("steps home from an alignment the ladder does not own", async () => {
    for (const exotic of [
      "Distributed",
      "Fill",
      "CenterAcrossSelection",
      "Justify",
    ]) {
      await boot();
      helpers.seed("Model!A1", [["Revenue"]]);
      helpers.select("Model!A1");
      helpers.setAlignment("Model!A1", exotic);

      await smt.applyAlignmentCycle();
      expect(helpers.cell("Model!A1").horizontalAlignment).toBe("Left");
    }
  });

  // Excel answers "SingleAccountant" / "DoubleAccountant" for the accounting
  // underlines, which draw the same rule as ours: same rung, next step.
  it("reads an accounting underline as the rule it draws", async () => {
    for (const [given, expected] of [
      ["SingleAccountant", "Double"],
      ["DoubleAccountant", "None"],
      ["Single", "Double"],
      ["None", "Single"],
    ]) {
      await boot();
      helpers.seed("Model!A1", [["Revenue"]]);
      helpers.select("Model!A1");
      helpers.setFont("Model!A1", { underline: given });

      await smt.applyUnderlineCycle();
      expect(helpers.font("Model!A1").underline).toBe(expected);
    }
  });

  it("treats a host answering null as no state at all", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    helpers.setIndent("Model!A1", null as unknown as number);
    helpers.setAlignment("Model!A1", null as unknown as string);
    helpers.setFont("Model!A1", { underline: null as unknown as string });

    await smt.applyIndentCycle();
    await smt.applyAlignmentCycle();
    await smt.applyUnderlineCycle();

    expect(helpers.cell("Model!A1").indentLevel).toBe(1);
    expect(helpers.cell("Model!A1").horizontalAlignment).toBe("Left");
    expect(helpers.font("Model!A1").underline).toBe("Single");
  });
});

describe("the wrong thing selected", () => {
  it("steps a single empty cell and leaves it empty", async () => {
    helpers.select("Model!D9");
    for (const [, run] of hygieneCycles()) await run();

    expect(helpers.cell("Model!D9").indentLevel).toBe(1);
    expect(helpers.cell("Model!D9").horizontalAlignment).toBe("Left");
    expect(helpers.font("Model!D9").underline).toBe("Single");
    expect(helpers.value("Model!D9")).toBe("");
    expect(helpers.formula("Model!D9")).toBe("");
  });

  // A merged block reads its format off any member cell, so the ladder still
  // lands on a rung and the whole block wears the same look.
  it("walks a merged block through the whole ladder", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:C1");
    helpers.select("Model!A1:C1");

    const walk: number[] = [];
    for (let press = 0; press < 5; press += 1) {
      await smt.applyIndentCycle();
      walk.push(helpers.cell("Model!A1").indentLevel);
    }
    expect(walk).toEqual([1, 2, 3, 0, 1]);
    expect(helpers.cell("Model!C1").indentLevel).toBe(1);
    expect(helpers.value("Model!A1")).toBe("Heading");
  });

  it("steps a selection whose cells hold errors, text and booleans", async () => {
    helpers.seed("Model!A1", [
      [{ formula: "=A99", value: "#REF!" }],
      [{ formula: "=NA()", value: "#N/A" }],
      ["text"],
      [true],
      [{ formula: '=IF(1=1,"","x")', value: "" }],
    ]);
    helpers.select("Model!A1:A5");
    for (const [, run] of hygieneCycles()) await run();

    // Nothing about the contents changed: a look is a look.
    expect(helpers.value("Model!A1")).toBe("#REF!");
    expect(helpers.formula("Model!A2")).toBe("=NA()");
    expect(helpers.value("Model!A4")).toBe(true);
    expect(helpers.cell("Model!A5").indentLevel).toBe(1);
  });

  it("writes into every area of a ctrl-clicked selection, whatever their shapes", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.merge("Model!C1:D1");
    helpers.selectAreas(["Model!A1", "Model!C1:D1", "Model!F1:F9"]);

    for (const [, run] of hygieneCycles()) await run();
    for (const cell of ["Model!A1", "Model!C1", "Model!F9"]) {
      expect(helpers.cell(cell).indentLevel).toBe(1);
      expect(helpers.cell(cell).horizontalAlignment).toBe("Left");
      expect(helpers.font(cell).underline).toBe("Single");
    }
  });

  // A selection far past the undo cap still steps - one property write is one
  // property write - but the net has to say it is off, and the whole stack
  // goes with it rather than offering a restore that is not safe.
  it("steps a 50 000-cell band with the undo net off and says so", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    await smt.applyIndentCycle();
    expect(smt.undoTarget()).not.toBeNull();

    // The band's own corner is A1, which the first press left on rung 1, so
    // the band steps to rung 2 - every cell of it, not just the corner.
    helpers.select("Model!A1:E10000");
    await smt.applyIndentCycle();

    expect(helpers.cell("Model!E10000").indentLevel).toBe(2);
    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(true);
  });

  // Ctrl+A: Excel answers cellCount -1 past 2^31-1 cells. The three hygiene
  // cycles write one scalar over the whole selection, exactly as the ribbon's
  // own Increase Indent does, so there is no cap to refuse them with - and the
  // fake host cannot run the write either (17 billion cells is its own guard,
  // not Excel's). What the pane owes here is proved on the real host: see the
  // launch-check row in the P2 report.
  it("counts a whole-sheet selection as over the cap, not under it", async () => {
    helpers.select("Model!A1:XFD1048576");
    const said = await rejects(() => smt.applyIndentCycle());

    // Not the pane's cap sentence and not a silent success: the rig's own
    // refusal, which proves the flow never tried to read the grid first.
    expect(said).toContain("17179869184-cell");
    expect(smt.undoTarget()).toBeNull();
  });
});

describe("a host that says no", () => {
  it("answers a protected sheet with the pane's sentence and changes nothing", async () => {
    for (const [stage, run] of hygieneCycles()) {
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

  // A protected sheet whose active cell is unlocked: the write is legal there
  // and refused a cell later, so the flow must not claim a clean refusal for a
  // block it did paint. The fake refuses the whole rectangle the way office.js
  // documents it, so a single-area selection is safe.
  it("changes nothing when the block holds one locked cell", async () => {
    helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
    helpers.select("Model!A1:A2");
    helpers.protectSheet("Model", ["Model!A1"]);

    expect(await rejects(() => smt.applyIndentCycle())).toBe(
      "Indent cycling: this sheet is protected, nothing was changed",
    );
    expect(helpers.cell("Model!A1").indentLevel).toBe(0);
    expect(helpers.cell("Model!A2").indentLevel).toBe(0);
  });

  // A ctrl-clicked selection on a partly unlocked sheet is several write
  // statements in one batch: the unlocked areas land, the locked one is
  // refused. protectedNote now words a batch of more than one area as partly
  // protected instead of promising an untouched sheet.
  it("does not promise an untouched sheet when it painted one area", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.seed("Model!C1", [["Cost"]]);
    helpers.selectAreas(["Model!A1", "Model!C1"]);
    helpers.protectSheet("Model", ["Model!A1"]);

    const said = await rejects(() => smt.applyIndentCycle());
    // The unlocked area was painted: the sentence may not deny it.
    expect(helpers.cell("Model!A1").indentLevel).toBe(1);
    expect(said).not.toContain("nothing was changed");
  });

  it("names the build when a ctrl-clicked selection meets a host below 1.9", async () => {
    await boot({
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.9"),
    });
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.selectAreas(["Model!A1", "Model!C1"]);

    for (const [stage, run] of hygieneCycles()) {
      expect(await rejects(run)).toBe(
        `${stage}: this Excel build can only act on one selected block.`,
      );
    }
  });

  it("still steps a single block on a host below 1.9", async () => {
    await boot({
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.9"),
    });
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1:A2");

    for (const [, run] of hygieneCycles()) await run();
    expect(helpers.cell("Model!A2").indentLevel).toBe(1);
  });

  it("leaves the cells alone when the host refuses the batch outright", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    helpers.failNextSync();

    await rejects(() => smt.applyIndentCycle());
    expect(helpers.cell("Model!A1").indentLevel).toBe(0);
  });
});

describe("pressed twice", () => {
  // Two presses that overlap read the same state and write the same rung, so
  // the workbook is consistent afterwards - one step, not a half-step and not
  // two different looks in one selection.
  it("lands every cell on the same rung when two presses overlap", async () => {
    helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
    helpers.select("Model!A1:A2");

    const both = await Promise.allSettled([
      smt.applyIndentCycle(),
      smt.applyIndentCycle(),
    ]);
    expect(both.map((entry) => entry.status)).toEqual([
      "fulfilled",
      "fulfilled",
    ]);
    expect(helpers.cell("Model!A1").indentLevel).toBe(1);
    expect(helpers.cell("Model!A2").indentLevel).toBe(1);
  });

  it("puts the whole selection back after a press, however many areas", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.seed("Model!C1", [["Cost"]]);
    helpers.selectAreas(["Model!A1", "Model!C1"]);
    helpers.setIndent("Model!A1", 2);
    helpers.setIndent("Model!C1", 2);

    await smt.applyIndentCycle();
    expect(helpers.cell("Model!C1").indentLevel).toBe(3);

    await smt.undoLastAction();
    expect(helpers.cell("Model!A1").indentLevel).toBe(2);
    expect(helpers.cell("Model!C1").indentLevel).toBe(2);
  });

  // The capture runs before the write, so a write the host refuses must not
  // spend a real slot - the ring is five deep, and the refusal would evict one
  // that landed. src/excel/undo.ts holds the entry pending until syncWrite
  // commits or discards it.
  it("spends no undo slot on a write the sheet refused", async () => {
    helpers.seed("Model!A1", [["a"], ["b"], ["c"], ["d"], ["e"], ["f"]]);
    for (const row of [1, 2, 3, 4, 5]) {
      helpers.select(`Model!A${String(row)}`);
      await smt.applyIndentCycle();
    }
    expect(smt.undoTarget()).toBe("Model!A5");

    helpers.select("Model!A6");
    helpers.protectSheet("Model");
    await rejects(() => smt.applyIndentCycle());

    expect(smt.undoTarget()).toBe("Model!A5");
  });
});
