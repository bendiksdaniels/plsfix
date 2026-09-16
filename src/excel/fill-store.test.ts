// Unit tests for originalFillColor, the pure lookup a table export uses to
// substitute a painted overlay's tint for the fill it covered: no Excel host
// needed, since a store's snapshot is plain in-memory state once painted.

import { describe, expect, it } from "vitest";
import { FillStore, originalFillColor } from "./fill-store";

const store = new FillStore("TEST_FILL_STORE", "the test overlay");

describe("originalFillColor", () => {
  it("is undefined when no store owns the cell", () => {
    expect(originalFillColor("no-store-sheet", 4, 2)).toBeUndefined();
  });

  it("is the colour a painted store remembers, null where it remembers no fill", () => {
    store.remember("sheet-a", "B4:C5", [
      ["none", "solid|#EEEEEE|#EEEEEE"],
      ["none", "none"],
    ]);
    expect(originalFillColor("sheet-a", 4, 2)).toBeNull(); // B4
    expect(originalFillColor("sheet-a", 4, 3)).toBe("#EEEEEE"); // C4
    expect(originalFillColor("sheet-a", 5, 3)).toBeNull(); // C5
  });

  it("is undefined outside the remembered rectangle, and on another sheet", () => {
    store.remember("sheet-b", "B4:C5", [
      ["none", "none"],
      ["none", "none"],
    ]);
    expect(originalFillColor("sheet-b", 6, 2)).toBeUndefined();
    expect(originalFillColor("sheet-z", 4, 2)).toBeUndefined();
  });
});
