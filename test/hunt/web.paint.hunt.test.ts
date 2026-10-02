// The paintbrush's borders on Excel for the web (rig 27.09): a slot captured
// on a cell without borders reads every edge as None, Thin, black, and the web
// draws a line for a None edge sent back with that weight and colour - so a
// paint must carry a None edge as its style alone, or it boxes every cell.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

const EDGES = ["top", "bottom", "left", "right"] as const;
const NO_LINE = { style: "None", weight: "Thin", color: "#000000" };
const RED_RULE = { style: "Continuous", weight: "Medium", color: "#C00000" };

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"], web: true }).helpers;
  smt = await import("../../src/excel");
});

function edgesOf(address: string): Record<string, unknown> {
  return Object.fromEntries(
    EDGES.map((edge) => [edge, helpers.border(address, edge)]),
  );
}

describe("the paintbrush's borders on the web", () => {
  it("draws no line from a slot captured on a cell without one", async () => {
    helpers.select("Model!E1");
    const plain = await smt.captureSlot(1);

    helpers.select("Model!A1:B2");
    await smt.applySlot(1, plain);

    for (const address of ["Model!A1", "Model!B1", "Model!A2", "Model!B2"]) {
      expect(edgesOf(address)).toEqual({
        top: NO_LINE,
        bottom: NO_LINE,
        left: NO_LINE,
        right: NO_LINE,
      });
    }
  });

  it("takes a box off the cell it paints, the way a format painter does", async () => {
    const boxed = helpers.sheet("Model").edit(0, 0);
    for (const edge of EDGES) boxed.borders[edge] = { ...RED_RULE };
    helpers.select("Model!E1");
    const plain = await smt.captureSlot(1);

    helpers.select("Model!A1");
    await smt.applySlot(1, plain);

    expect(edgesOf("Model!A1")).toEqual({
      top: NO_LINE,
      bottom: NO_LINE,
      left: NO_LINE,
      right: NO_LINE,
    });
  });

  it("paints the one rule a slot carries and nothing else", async () => {
    helpers.sheet("Model").edit(0, 4).borders.bottom = { ...RED_RULE };
    helpers.select("Model!E1");
    const ruled = await smt.captureSlot(2);

    helpers.select("Model!A1");
    await smt.applySlot(2, ruled);

    expect(edgesOf("Model!A1")).toEqual({
      top: NO_LINE,
      bottom: RED_RULE,
      left: NO_LINE,
      right: NO_LINE,
    });
  });
});
