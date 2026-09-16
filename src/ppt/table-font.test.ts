// Unit tests for uniformSize, the pure modal-font-size rule behind a table's
// insert-time uniformCellProperties: no PowerPoint host needed.

import { describe, expect, it } from "vitest";
import type { TableCell } from "../link/model";
import { uniformSize } from "./table-font";

const payload = (cells: TableCell[][]) =>
  ({ cells }) as Parameters<typeof uniformSize>[0];

describe("uniformSize", () => {
  it("is undefined when no cell names a size", () => {
    expect(uniformSize(payload([[{ t: "a" }, { t: "b" }]]))).toBeUndefined();
  });

  it("is the one size present even when a single cell names it", () => {
    expect(uniformSize(payload([[{ t: "a", z: 11 }, { t: "b" }]]))).toBe(11);
  });

  it("is the size most cells carry", () => {
    const cells: TableCell[][] = [
      [
        { t: "a", z: 11 },
        { t: "b", z: 11 },
      ],
      [
        { t: "c", z: 11 },
        { t: "d", z: 9 },
      ],
    ];
    expect(uniformSize(payload(cells))).toBe(11);
  });

  it("keeps the smaller size on a tie", () => {
    const cells: TableCell[][] = [
      [
        { t: "a", z: 11 },
        { t: "b", z: 9 },
      ],
    ];
    expect(uniformSize(payload(cells))).toBe(9);
  });
});
