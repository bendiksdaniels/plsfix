import { describe, expect, it } from "vitest";
import {
  emptySlots,
  PAINT_SLOT_COUNT,
  type PaintSlot,
  parseSlots,
  serializeSlots,
  slotLabel,
} from "./paintbrush";

const SLOT: PaintSlot = {
  numberFormat: "#,##0.0;[Red](#,##0.0);-",
  font: {
    name: "Aptos",
    size: 10,
    bold: true,
    italic: false,
    color: "#1F1D1B",
  },
  fill: "#F2E8DF",
  horizontalAlignment: "Right",
  borders: {
    top: { style: "None", weight: "Thin", color: "#000000" },
    bottom: { style: "Continuous", weight: "Medium", color: "#B27E54" },
    left: { style: "None", weight: "Thin", color: "#000000" },
    right: { style: "None", weight: "Thin", color: "#000000" },
  },
};

function stored(slot: unknown): string {
  return JSON.stringify([slot]);
}

describe("paint slot codec", () => {
  it("round trips captured slots", () => {
    const slots = emptySlots();
    slots[0] = SLOT;
    slots[2] = { ...SLOT, numberFormat: "0.0%", fill: null };

    expect(parseSlots(serializeSlots(slots))).toEqual(slots);
  });

  it("keeps an unfilled slot apart from a coloured one", () => {
    const parsed = parseSlots(serializeSlots([{ ...SLOT, fill: null }]));

    expect(parsed[0]?.fill).toBeNull();
    expect(parseSlots(serializeSlots([SLOT]))[0]?.fill).toBe("#F2E8DF");
  });

  it("is always three slots long, whatever it read", () => {
    expect(emptySlots()).toEqual([null, null, null]);
    expect(parseSlots(serializeSlots([SLOT]))).toHaveLength(PAINT_SLOT_COUNT);
    expect(parseSlots(JSON.stringify([SLOT, SLOT, SLOT, SLOT]))).toHaveLength(
      PAINT_SLOT_COUNT,
    );
  });

  it("reads garbage as empty slots", () => {
    for (const raw of [null, "", "not json", "{}", "[1,2,3]", "[null]"]) {
      expect(parseSlots(raw)).toEqual(emptySlots());
    }
  });

  it("refuses a slot that is missing or mistyping a field", () => {
    const cases: unknown[] = [
      { ...SLOT, numberFormat: 5 },
      { ...SLOT, horizontalAlignment: undefined },
      { ...SLOT, fill: 0 },
      { ...SLOT, font: { ...SLOT.font, size: "10" } },
      { ...SLOT, font: { ...SLOT.font, bold: "yes" } },
      { ...SLOT, font: { ...SLOT.font, name: "" } },
      { ...SLOT, borders: { top: SLOT.borders.top } },
      { ...SLOT, borders: { ...SLOT.borders, left: { style: "None" } } },
    ];

    for (const broken of cases) {
      expect(parseSlots(stored(broken))).toEqual(emptySlots());
    }
  });

  it("labels a slot by its number format and font", () => {
    expect(slotLabel(SLOT)).toBe("#,##0.0;[Red](#,##0.0);- · Aptos");
    expect(slotLabel(null)).toBe("empty");
  });
});
