// Unit tests for headerRow, the pure rule behind the `h` payload flag: no
// Excel host needed, since it only ever looks at the cells already read.

import { describe, expect, it } from "vitest";
import type { TableCell } from "../link/model";
import { headerRow } from "./link-table";

function cell(t: string, b?: true): TableCell {
  return b === true ? { t, b } : { t };
}

describe("headerRow", () => {
  it("reads a header when every non-empty cell of row 0 is bold", () => {
    const cells: TableCell[][] = [
      [cell("Revenue", true), cell("Costs", true)],
      [cell("1 000"), cell("-400")],
    ];
    expect(headerRow(cells)).toBe(true);
  });

  it("refuses a row with one non-bold non-empty cell", () => {
    const cells: TableCell[][] = [
      [cell("Revenue", true), cell("Costs")],
      [cell("1 000"), cell("-400")],
    ];
    expect(headerRow(cells)).toBe(false);
  });

  it("refuses a single-row table: there is no body to be a header over", () => {
    const cells: TableCell[][] = [[cell("Revenue", true), cell("Costs", true)]];
    expect(headerRow(cells)).toBe(false);
  });

  it("refuses a table whose row 0 is entirely empty text", () => {
    const cells: TableCell[][] = [
      [cell(""), cell("  ")],
      [cell("1 000"), cell("-400")],
    ];
    expect(headerRow(cells)).toBe(false);
  });

  it("ignores a bold empty cell sitting beside the labelled ones", () => {
    const cells: TableCell[][] = [
      [cell("Revenue", true), cell("", true), cell("Costs", true)],
      [cell("1 000"), cell(""), cell("-400")],
    ];
    expect(headerRow(cells)).toBe(true);
  });

  it("reads a header off numbers-only cells, bold is the only rule", () => {
    const cells: TableCell[][] = [
      [cell("2024", true), cell("2025", true)],
      [cell("1 000"), cell("1 200")],
    ];
    expect(headerRow(cells)).toBe(true);
  });
});
