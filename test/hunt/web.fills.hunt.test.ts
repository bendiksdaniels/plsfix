// Excel for the web reads a solid fill back with a null pattern and an unfilled
// cell with colour "" (rig 27.09): every flow that snapshots or captures a fill
// must still hand the modeller's own fills back - the audit overlay, the
// linked-cell highlight, the pinstripes' second press and the paintbrush.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeCell,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import { FakeRelay } from "../fakerelay";
import type * as ExcelModule from "../../src/excel";
import { createWorkspace, type KeyStore } from "../../src/link/workspace";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

const YELLOW = { color: "#FFFF00", pattern: "Solid", patternColor: "#FFFFFF" };
// The web's solid white reads exactly like Excel for Mac's unfilled cell.
const WHITE = { color: "#FFFFFF", pattern: "Solid", patternColor: "#FFFFFF" };
const STRIPED = {
  color: "#EEDDCC",
  pattern: "LightUp",
  patternColor: "#0057B8",
};

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

// A block the overlay stripes (a formula in every row) wearing three of the
// modeller's own fills and one cell with none (C3).
function seedModel(): void {
  const formula = (r1c1: string) => ({ formula: r1c1, r1c1, value: 1 });
  helpers.seed("Model!A1", [
    [formula("=R[1]C"), formula("=R[1]C"), formula("=R[1]C")],
    [formula("=R[1]C"), formula("=RC[-1]"), 5],
    [formula("=X"), "", ""],
  ]);
  helpers.setFill("Model!A1", YELLOW);
  helpers.setFill("Model!B2", WHITE);
  helpers.setFill("Model!C1", STRIPED);
}

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

describe("the audit overlay on the web", () => {
  it("hands back a solid, a white, a striped and an empty cell exactly", async () => {
    seedModel();
    const before = shown("Model");
    helpers.select("Model!A1:C3");

    expect(await smt.toggleAuditOverlay()).toBe(true);
    expect(helpers.fill("Model!A1").pattern).not.toBe("Solid");
    expect(await smt.toggleAuditOverlay()).toBe(false);

    expect(shown("Model")).toEqual(before);
  });
});

describe("the linked-cell highlight on the web", () => {
  it("hands back the fills under the linked range exactly", async () => {
    seedModel();
    const ws = await createWorkspace(memoryStore());
    helpers.select("Model!A1:C3");
    await smt.exportSelection(ws, new FakeRelay());
    const before = shown("Model");

    expect(await smt.toggleLinkHighlight()).toBe(true);
    expect(await smt.toggleLinkHighlight()).toBe(false);

    expect(shown("Model")).toEqual(before);
  });
});

describe("the pinstripes on the web", () => {
  it("clears its own bands on the second press", async () => {
    helpers.seed(
      "Model!A1",
      Array.from({ length: 5 }, (_unused, row) => [row, row + 1, row + 2]),
    );
    helpers.select("Model!A1:C5");

    expect(await smt.applyPinstripes("rows")).toBe("Pinstripes: 2 rows banded");
    expect(await smt.applyPinstripes("rows")).toBe(
      "Pinstripes: 2 rows cleared",
    );
    expect(helpers.fill("Model!A2").pattern).toBe("None");
    expect(helpers.fill("Model!C4").pattern).toBe("None");
  });
});

describe("the paintbrush on the web", () => {
  it("captures an unfilled cell as no fill and clears the fill it paints over", async () => {
    helpers.select("Model!E1");
    const slot = await smt.captureSlot(1);
    expect(slot.fill).toBeNull();

    helpers.setFill("Model!A1:B1", YELLOW);
    helpers.select("Model!A1:B1");
    await smt.applySlot(1, slot);

    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.fill("Model!B1").pattern).toBe("None");
  });

  it("captures a solid fill by its colour, white included", async () => {
    helpers.setFill("Model!E1", WHITE);
    helpers.select("Model!E1");
    const white = await smt.captureSlot(2);
    expect(white.fill).toBe("#FFFFFF");

    helpers.select("Model!A1");
    await smt.applySlot(2, white);
    expect(helpers.fill("Model!A1")).toMatchObject({
      pattern: "Solid",
      color: "#FFFFFF",
    });
  });
});
