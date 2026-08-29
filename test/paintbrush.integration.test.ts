// The paintbrush against the strict fake host: capture reads the active cell
// and exactly the fields a slot carries, apply paints every cell of the
// selection (interior borders included), SMT Undo puts the old look back, and
// an empty slot is refused by name before Excel is touched.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type { PaintSlot } from "../src/paintbrush";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
});

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

const PLAIN_BORDER = { style: "None", weight: "Thin", color: "#000000" };

// A source cell wearing something on every field a slot carries.
function paintSource(address: string): void {
  helpers.setFont(address, {
    name: "Aptos",
    size: 12,
    bold: true,
    italic: true,
    color: "#1F1D1B",
  });
  helpers.setFill(address, { color: "#F2E8DF", pattern: "Solid" });
  helpers.setNumberFormat(address, "#,##0.0");
  const cell = helpers.sheet("Model").edit(0, 0);
  cell.horizontalAlignment = "Right";
  cell.borders.bottom = {
    style: "Continuous",
    color: "#B27E54",
    weight: "Medium",
  };
}

async function captureFromSource(index = 1): Promise<PaintSlot> {
  paintSource("Model!A1");
  helpers.select("Model!A1");
  return smt.captureSlot(index);
}

function expectPainted(address: string): void {
  expect(helpers.numberFormat(address)).toBe("#,##0.0");
  expect(helpers.font(address)).toMatchObject({
    name: "Aptos",
    size: 12,
    bold: true,
    italic: true,
    color: "#1F1D1B",
  });
  expect(helpers.fill(address).color).toBe("#F2E8DF");
  expect(helpers.cell(address).horizontalAlignment).toBe("Right");
  expect(helpers.border(address, "bottom")).toEqual({
    style: "Continuous",
    color: "#B27E54",
    weight: "Medium",
  });
  expect(helpers.border(address, "top")).toEqual(PLAIN_BORDER);
}

describe("paintbrush capture", () => {
  it("reads exactly the fields a slot carries", async () => {
    expect(await captureFromSource()).toEqual({
      numberFormat: "#,##0.0",
      font: {
        name: "Aptos",
        size: 12,
        bold: true,
        italic: true,
        color: "#1F1D1B",
      },
      fill: "#F2E8DF",
      horizontalAlignment: "Right",
      borders: {
        top: PLAIN_BORDER,
        bottom: { style: "Continuous", weight: "Medium", color: "#B27E54" },
        left: PLAIN_BORDER,
        right: PLAIN_BORDER,
      },
    } satisfies PaintSlot);
  });

  it("reads the active cell, not the first cell of the selection", async () => {
    paintSource("Model!A1");
    helpers.select("Model!A1:B2");
    helpers.setActiveCell("Model!B2");

    const slot = await smt.captureSlot(1);
    expect(slot.numberFormat).toBe("General");
    expect(slot.font.name).toBe("Calibri");
    expect(slot.fill).toBeNull();
  });

  it("refuses a slot number nobody has", async () => {
    helpers.select("Model!A1");
    expect(await rejects(() => smt.captureSlot(4))).toBe(
      "paintbrush: slot 4 does not exist",
    );
  });
});

describe("paintbrush apply", () => {
  it("paints every cell of a 2x2 selection", async () => {
    const slot = await captureFromSource(2);
    helpers.seed("Data!B2", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Data!B2:C3");
    await smt.applySlot(2, slot);

    for (const address of ["Data!B2", "Data!C2", "Data!B3", "Data!C3"]) {
      expectPainted(address);
    }
    // Formatting only: the numbers underneath are the modeller's.
    expect(helpers.value("Data!C3")).toBe(4);
  });

  it("clears a fill the slot does not carry", async () => {
    helpers.select("Model!D1");
    const empty = await smt.captureSlot(3);
    helpers.setFill("Data!A1", { color: "#123456", pattern: "Solid" });
    helpers.select("Data!A1");
    await smt.applySlot(3, empty);

    expect(helpers.fill("Data!A1")).toEqual({
      color: "#FFFFFF",
      pattern: "None",
      patternColor: "#FFFFFF",
    });
  });

  it("puts the old look back on undo", async () => {
    const slot = await captureFromSource();
    helpers.setFill("Data!A1:B2", { color: "#123456", pattern: "Solid" });
    helpers.setNumberFormat("Data!A1:B2", "0.00%");
    const before = helpers.cellMap("Data");

    helpers.select("Data!A1:B2");
    await smt.applySlot(1, slot);
    expect(helpers.cellMap("Data")).not.toEqual(before);

    await smt.undoLastAction();
    expect(helpers.cellMap("Data")).toEqual(before);
  });

  it("refuses an empty slot by name", async () => {
    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.applySlot(2, null))).toBe(
      "paintbrush: slot 2 is empty",
    );
  });

  it("refuses a ctrl-clicked selection and an oversized one", async () => {
    const slot = await captureFromSource();

    helpers.selectAreas(["Data!A1:B2", "Data!D1:D4"]);
    expect(await rejects(() => smt.applySlot(1, slot))).toBe(
      "paintbrush: select a single range",
    );

    helpers.select("Data!A1:A5001");
    expect(await rejects(() => smt.applySlot(1, slot))).toBe(
      "Paintbrush supports up to 5,000 selected cells at once.",
    );
  });
});
