// The slot codec against storage it did not write. Local storage is shared
// with whatever else the webview kept, so every half-shaped slot has to read
// as empty rather than as a slot that paints three of its four border rules.
import { describe, expect, it } from "vitest";
import { emptySlots, parseSlots, type PaintSlot } from "./paintbrush";

const BORDER = { style: "None", weight: "Thin", color: "#000000" };

const SLOT: PaintSlot = {
  numberFormat: "#,##0",
  font: {
    name: "Aptos",
    size: 10,
    bold: false,
    italic: false,
    color: "#1F1D1B",
  },
  fill: null,
  horizontalAlignment: "General",
  borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
};

function first(slot: unknown): PaintSlot | null {
  return parseSlots(JSON.stringify([slot]))[0] ?? null;
}

describe("a slot storage did not write", () => {
  it("keeps the well-formed one it was given", () => {
    expect(first(SLOT)).toEqual(SLOT);
  });

  it("reads a non-array and a non-object entry as empty", () => {
    expect(parseSlots(JSON.stringify({ slot: SLOT }))).toEqual(emptySlots());
    expect(parseSlots(JSON.stringify([1, "two", null]))).toEqual(emptySlots());
  });

  it("refuses a slot whose own fields are the wrong type", () => {
    expect(first({ ...SLOT, numberFormat: 3 })).toBeNull();
    expect(first({ ...SLOT, horizontalAlignment: null })).toBeNull();
    expect(first({ ...SLOT, fill: 16711680 })).toBeNull();
  });

  it("refuses a font that could not be written back", () => {
    expect(first({ ...SLOT, font: null })).toBeNull();
    expect(first({ ...SLOT, font: { ...SLOT.font, name: "" } })).toBeNull();
    expect(first({ ...SLOT, font: { ...SLOT.font, size: 0 } })).toBeNull();
    expect(
      first({ ...SLOT, font: { ...SLOT.font, size: Number.NaN } }),
    ).toBeNull();
    expect(first({ ...SLOT, font: { ...SLOT.font, bold: "yes" } })).toBeNull();
    expect(first({ ...SLOT, font: { ...SLOT.font, italic: 1 } })).toBeNull();
    expect(first({ ...SLOT, font: { ...SLOT.font, color: "" } })).toBeNull();
  });

  // All four edges or nothing: a half-read border set would paint one rule and
  // leave the other three as they were.
  it("refuses a border set with one edge missing or malformed", () => {
    const borders = SLOT.borders;
    expect(first({ ...SLOT, borders: null })).toBeNull();
    expect(
      first({ ...SLOT, borders: { ...borders, right: undefined } }),
    ).toBeNull();
    expect(
      first({
        ...SLOT,
        borders: { ...borders, left: { ...BORDER, style: "" } },
      }),
    ).toBeNull();
    expect(
      first({
        ...SLOT,
        borders: { ...borders, bottom: { ...BORDER, weight: 2 } },
      }),
    ).toBeNull();
    expect(
      first({ ...SLOT, borders: { ...borders, top: { ...BORDER, color: 0 } } }),
    ).toBeNull();
  });
});
