import { describe, expect, it } from "vitest";
import { TAG_VALUE_MAX, type TableCell } from "./model";
import { decodePaintMap, encodePaintMap, paintKey } from "./paint-map";

function cell(f?: string): TableCell {
  return f === undefined ? { t: "x" } : { t: "x", f };
}

describe("encodePaintMap / decodePaintMap", () => {
  it("round-trips filled cells, folding a run into a range and keeping a gap apart", () => {
    const cells: TableCell[][] = [
      [cell("#fff"), cell("#fff"), cell("#fff"), cell(), cell(), cell("#fff")],
      [cell(), cell(), cell(), cell()],
      [cell(), cell(), cell(), cell("#fff")],
    ];
    const encoded = encodePaintMap(cells);
    expect(encoded).toBe("1;0:0-2,5;2:3");
    expect(decodePaintMap(encoded)).toEqual(
      new Set([
        paintKey(0, 0),
        paintKey(0, 1),
        paintKey(0, 2),
        paintKey(0, 5),
        paintKey(2, 3),
      ]),
    );
  });

  it("encodes a grid with no filled cell as the bare version, decoding to an empty set", () => {
    const cells: TableCell[][] = [
      [cell(), cell()],
      [cell(), cell()],
    ];
    expect(encodePaintMap(cells)).toBe("1;");
    expect(decodePaintMap("1;")).toEqual(new Set());
  });

  it("reads absent, malformed or a foreign version as legacy (null)", () => {
    expect(decodePaintMap(undefined)).toBeNull();
    expect(decodePaintMap(null)).toBeNull();
    expect(decodePaintMap("garbage")).toBeNull();
    expect(decodePaintMap("1;no-colon-here")).toBeNull();
    expect(decodePaintMap("1;0:not-a-number")).toBeNull();
    expect(decodePaintMap("1;0:2-1")).toBeNull();
    expect(decodePaintMap("2;0:0")).toBeNull();
  });

  it("throws once the encoded map would exceed a tag value", () => {
    // 80 rows of 30 non-adjacent fills apiece: no run collapses, so the
    // string is comfortably past TAG_VALUE_MAX.
    const cells: TableCell[][] = Array.from({ length: 80 }, () =>
      Array.from({ length: 60 }, (_unused, c) =>
        cell(c % 2 === 0 ? "#fff" : undefined),
      ),
    );
    expect(() => encodePaintMap(cells)).toThrow(
      new RegExp(`exceeds ${String(TAG_VALUE_MAX)}`),
    );
  });
});
