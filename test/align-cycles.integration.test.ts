// The three hygiene cycles - indent, horizontal alignment, font underline -
// against the fake host: the position is read off the active cell and the next
// look is written into every area of the selection. Plus the universality rows
// every Excel tool answers for: a ctrl-clicked selection, merged cells, a
// protected sheet, a selection over the undo cap and an empty one.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
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

function seedBlock(): void {
  helpers.seed("Model!A1", [["Revenue"], ["Cost"], ["Profit"]]);
  helpers.select("Model!A1:A3");
}

// Two blocks a modeller would ctrl-click: a label column and its total.
function selectTwoBlocks(): void {
  helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
  helpers.seed("Model!C1", [["Total"], ["Margin"]]);
  helpers.selectAreas(["Model!A1:A2", "Model!C1:C2"]);
}

beforeEach(async () => {
  await boot();
});

describe("the indent cycle", () => {
  it("steps 0, 1, 2, 3 and back home", async () => {
    seedBlock();
    const walk: number[] = [];
    for (let press = 0; press < 4; press += 1) {
      await smt.applyIndentCycle();
      walk.push(helpers.cell("Model!A1").indentLevel);
    }
    expect(walk).toEqual([1, 2, 3, 0]);
  });

  it("indents every cell of the selection, not just the first", async () => {
    seedBlock();
    await smt.applyIndentCycle();
    expect(helpers.cell("Model!A3").indentLevel).toBe(1);
  });

  it("reads the position off the active cell of a hand-set block", async () => {
    seedBlock();
    helpers.setIndent("Model!A1:A3", 2);
    await smt.applyIndentCycle();
    expect(helpers.cell("Model!A1").indentLevel).toBe(3);
  });

  it("captures the previous indent, so Undo puts it back", async () => {
    seedBlock();
    helpers.setIndent("Model!A1:A3", 1);
    await smt.applyIndentCycle();
    expect(helpers.cell("Model!A2").indentLevel).toBe(2);

    await smt.undoLastAction();
    expect(helpers.cell("Model!A2").indentLevel).toBe(1);
  });
});

describe("the alignment cycle", () => {
  it("steps Left, Center, Right, then General", async () => {
    seedBlock();
    const walk: string[] = [];
    for (let press = 0; press < 4; press += 1) {
      await smt.applyAlignmentCycle();
      walk.push(helpers.cell("Model!A1").horizontalAlignment);
    }
    expect(walk).toEqual(["Left", "Center", "Right", "General"]);
  });

  it("starts at Left from an alignment the cycle does not own", async () => {
    seedBlock();
    helpers.setAlignment("Model!A1:A3", "Justify");
    await smt.applyAlignmentCycle();
    expect(helpers.cell("Model!A1").horizontalAlignment).toBe("Left");
  });

  it("aligns every cell of the selection", async () => {
    seedBlock();
    await smt.applyAlignmentCycle();
    expect(helpers.cell("Model!A3").horizontalAlignment).toBe("Left");
  });
});

describe("the underline cycle", () => {
  it("steps Single, Double, then off", async () => {
    seedBlock();
    const walk: string[] = [];
    for (let press = 0; press < 3; press += 1) {
      await smt.applyUnderlineCycle();
      walk.push(helpers.font("Model!A1").underline);
    }
    expect(walk).toEqual(["Single", "Double", "None"]);
  });

  it("counts Excel's accounting underline as the single it draws", async () => {
    seedBlock();
    helpers.setFont("Model!A1:A3", { underline: "SingleAccountant" });
    await smt.applyUnderlineCycle();
    expect(helpers.font("Model!A1").underline).toBe("Double");
  });

  it("underlines every cell of the selection", async () => {
    seedBlock();
    await smt.applyUnderlineCycle();
    expect(helpers.font("Model!A3").underline).toBe("Single");
  });
});

describe("the universality rows", () => {
  it("writes into every area of a ctrl-clicked selection", async () => {
    selectTwoBlocks();
    await smt.applyIndentCycle();
    await smt.applyAlignmentCycle();
    await smt.applyUnderlineCycle();

    expect(helpers.cell("Model!C2").indentLevel).toBe(1);
    expect(helpers.cell("Model!C2").horizontalAlignment).toBe("Left");
    expect(helpers.font("Model!C2").underline).toBe("Single");
  });

  it("formats a merged block, which Excel allows", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A1:B1");
    helpers.select("Model!A1:B1");

    await smt.applyAlignmentCycle();
    expect(helpers.cell("Model!A1").horizontalAlignment).toBe("Left");
  });

  it("answers a protected sheet with the pane's own sentence", async () => {
    seedBlock();
    helpers.protectSheet("Model");
    const before = helpers.cellMap("Model");

    expect(await rejects(() => smt.applyIndentCycle())).toBe(
      "Indent cycling: this sheet is protected, nothing was changed",
    );
    expect(await rejects(() => smt.applyAlignmentCycle())).toBe(
      "Alignment cycling: this sheet is protected, nothing was changed",
    );
    expect(await rejects(() => smt.applyUnderlineCycle())).toBe(
      "Underline cycling: this sheet is protected, nothing was changed",
    );
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("still steps a selection past the undo cap, and says so", async () => {
    helpers.select("Model!A1:A6000");
    await smt.applyIndentCycle();

    expect(helpers.cell("Model!A6000").indentLevel).toBe(1);
    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(true);
  });

  it("steps an empty selection: a look is a look with no values in it", async () => {
    helpers.select("Model!D4:E5");
    await smt.applyUnderlineCycle();

    expect(helpers.font("Model!D4").underline).toBe("Single");
    expect(helpers.font("Model!E5").underline).toBe("Single");
    expect(helpers.value("Model!D4")).toBe("");
  });
});
