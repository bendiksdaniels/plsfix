// pls,fix Undo on Excel for the web (rig 27.09): the captured cells carry an
// @odata.type annotation on every level, the borders collection included, and
// a None edge sent back with its weight or colour is drawn as a line. A preset,
// a format cycle and an autocolor pass, each undone, must restore the block.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeCell,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"], web: true }).helpers;
  smt = await import("../../src/excel");
});

// What the web shows of a fill: a solid one is its colour, whatever pattern
// colour hides under it (the web never reads one back).
function shown(sheet: string): Record<string, unknown> {
  const cells: Record<string, FakeCell> = helpers.cellMap(sheet);
  return Object.fromEntries(
    Object.entries(cells).map(([key, cell]) => {
      const { fill } = cell;
      const seen =
        fill.pattern === "Solid"
          ? { pattern: "Solid", color: fill.color }
          : fill.pattern === "None"
            ? { pattern: "None" }
            : fill;
      return [key, { ...cell, fill: seen }];
    }),
  );
}

// Numbers, formulas and a label wearing a solid fill, a striped one, a bold
// coloured font and a thick red rule under A2.
function seedModel(): void {
  helpers.seed("Model!A1", [
    [100, { formula: "=A1*2", value: 200 }, "Revenue"],
    [{ formula: "=A1+B1", value: 300 }, 50, 7],
  ]);
  helpers.setFill("Model!A1", { color: "#FFFF00", pattern: "Solid" });
  helpers.setFill("Model!B1", {
    color: "#EEDDCC",
    pattern: "LightUp",
    patternColor: "#0057B8",
  });
  helpers.setFont("Model!C2", { bold: true, color: "#123456" });
  helpers.sheet("Model").edit(1, 0).borders.bottom = {
    style: "Continuous",
    weight: "Medium",
    color: "#C00000",
  };
  helpers.select("Model!A1:C2");
}

async function pressThenUndo(press: () => Promise<unknown>): Promise<void> {
  seedModel();
  const before = shown("Model");
  await press();
  expect(shown("Model")).not.toEqual(before);

  expect(await smt.undoLastAction()).toBe(
    "Undone: Model!A1:C2. Nothing more to undo.",
  );
  expect(shown("Model")).toEqual(before);
}

describe("pls,fix Undo on the web", () => {
  it("restores a Header preset exactly", async () => {
    await pressThenUndo(() => smt.applyPreset("header"));
  });

  it("restores a border cycle exactly", async () => {
    await pressThenUndo(() => smt.applyBorderCycle());
  });

  it("restores an autocolor pass exactly", async () => {
    await pressThenUndo(() => smt.autocolorSelection());
  });
});
