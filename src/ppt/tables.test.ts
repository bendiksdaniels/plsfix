import { describe, expect, it } from "vitest";

import { columnWidths } from "./tables";

const payload = (widths: number[]) =>
  ({ widths, cols: widths.length }) as Parameters<typeof columnWidths>[0];

describe("columnWidths", () => {
  it("scales the source widths to the table width in whole points", () => {
    const widths = columnWidths(payload([322, 49, 77, 77]), 400);
    expect(widths.every((one) => Number.isInteger(one))).toBe(true);
    expect(widths.reduce((sum, one) => sum + one, 0)).toBe(400);
    expect(widths[0]).toBeGreaterThan(widths[1]!);
  });

  it("keeps widths that already fit and never drops a column to nothing", () => {
    expect(columnWidths(payload([96, 48]), 144)).toEqual([96, 48]);
    expect(columnWidths(payload([1000, 1]), 200)).toEqual([199, 1]);
  });
});
