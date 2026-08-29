import { describe, expect, it } from "vitest";
import {
  anchorName,
  decodePayload,
  decodeRegistry,
  decodeTag,
  encodePayload,
  encodeRegistry,
  tryDecodeRegistry,
  encodeTag,
  isLinkId,
  newLinkId,
  overTableCap,
  payloadBytes,
  sourceLabel,
  TABLE_MAX_COLS,
  TABLE_MAX_ROWS,
  type LinkTag,
  type Payload,
  type TablePayload,
} from "./model";

const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_0123abcd",
};
const tag: LinkTag = {
  v: 1,
  id: "0123abcd".repeat(4),
  kind: "range",
  rev: 3,
  src,
  pushedAt: "2026-08-28T19:00:00.000Z",
};

describe("link ids", () => {
  it("are 32 lowercase hex chars from 16 random bytes", () => {
    const id = newLinkId((n) => new Uint8Array(n).fill(0xab));
    expect(id).toBe("ab".repeat(16));
    expect(isLinkId(id)).toBe(true);
    expect(isLinkId("AB".repeat(16))).toBe(false);
    expect(isLinkId("abc")).toBe(false);
  });
  it("anchor names are valid Excel defined names", () => {
    expect(anchorName("0123abcd".repeat(4))).toBe("PLSFIX_LINK_0123abcd");
  });
});

describe("tags", () => {
  it("round-trip and reject garbage", () => {
    expect(decodeTag(encodeTag(tag))).toEqual(tag);
    expect(decodeTag("{}")).toBeNull();
    expect(decodeTag("not json")).toBeNull();
    expect(decodeTag(null)).toBeNull();
    expect(decodeTag(JSON.stringify({ ...tag, v: 2 }))).toBeNull();
  });
  it("refuse oversize values", () => {
    expect(() =>
      encodeTag({ ...tag, src: { ...src, workbook: "x".repeat(3000) } }),
    ).toThrow(/2048/);
  });
});

describe("registry", () => {
  it("round-trips and defaults to empty", () => {
    const registry = {
      v: 1 as const,
      links: [
        {
          id: tag.id,
          kind: "range" as const,
          anchor: "PLSFIX_LINK_0123abcd",
          label: "Revenue",
          token: "t".repeat(43),
          createdAt: tag.pushedAt,
          lastPushedAt: null,
          rev: 0,
        },
      ],
    };
    expect(decodeRegistry(encodeRegistry(registry))).toEqual(registry);
    expect(decodeRegistry(undefined)).toEqual({ v: 1, links: [] });
    expect(decodeRegistry("[1,2]")).toEqual({ v: 1, links: [] });
  });

  it("tells an unreadable registry from an absent one", () => {
    const empty = { v: 1 as const, links: [] };
    expect(tryDecodeRegistry(encodeRegistry(empty))).toEqual(empty);
    expect(tryDecodeRegistry(null)).toBeNull();
    expect(tryDecodeRegistry(undefined)).toBeNull();
    expect(tryDecodeRegistry("{not json")).toBeNull();
    expect(tryDecodeRegistry("[1,2]")).toBeNull();
    // A newer schema is readable JSON we still must not overwrite.
    expect(tryDecodeRegistry('{"v":2,"links":[]}')).toBeNull();
    expect(tryDecodeRegistry('{"v":1,"links":[{"id":"x"}]}')).toBeNull();
  });
});

describe("payload", () => {
  it("round-trips through UTF-8 bytes and weighs its base64", () => {
    const payload: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 800,
      height: 400,
      png: "iVBORw0KGgo=",
      src,
      pushedAt: tag.pushedAt,
      hash: "ab".repeat(32),
    };
    expect(decodePayload(encodePayload(payload))).toEqual(payload);
    expect(payloadBytes(payload)).toBe(payload.png.length);
    expect(() => decodePayload(new TextEncoder().encode('{"v":1}'))).toThrow(
      /payload/,
    );
  });
});

describe("table payload", () => {
  const table: TablePayload = {
    v: 1,
    kind: "table",
    rows: 2,
    cols: 2,
    cells: [
      [{ t: "Revenue" }, { t: "1 000", b: true, a: "r" }],
      [
        { t: "Costs", i: true, c: "#FF0000", z: 9 },
        { t: "-400", f: "#EEEEEE", a: "c" },
      ],
    ],
    widths: [80, 60],
    src,
    pushedAt: tag.pushedAt,
    hash: "cd".repeat(32),
  };

  it("round-trips every cell field through UTF-8 bytes", () => {
    expect(decodePayload(encodePayload(table))).toEqual(table);
  });

  it("refuses a payload whose fields are the wrong shape", () => {
    const bad = (patch: Record<string, unknown>): (() => Payload) => {
      const broken = { ...table, ...patch };
      return () =>
        decodePayload(new TextEncoder().encode(JSON.stringify(broken)));
    };
    expect(bad({ rows: "2" })).toThrow(/payload/);
    expect(bad({ cols: null })).toThrow(/payload/);
    expect(bad({ widths: [80, "60"] })).toThrow(/payload/);
    expect(bad({ cells: [[{ t: 1 }]] })).toThrow(/payload/);
    expect(bad({ cells: [[{ t: "x", b: false }]] })).toThrow(/payload/);
    expect(bad({ cells: [[{ t: "x", a: "middle" }]] })).toThrow(/payload/);
    expect(bad({ cells: [[{ t: "x", z: "9" }]] })).toThrow(/payload/);
    expect(bad({ cells: [["x"]] })).toThrow(/payload/);
    // The row count is the grid the cells make, not a number beside it.
    expect(bad({ rows: 3 })).toThrow(/payload/);
    expect(bad({ cols: 3 })).toThrow(/payload/);
    expect(bad({ widths: [80] })).toThrow(/payload/);
  });

  it("caps a table at 60 rows and 20 columns", () => {
    expect(TABLE_MAX_ROWS).toBe(60);
    expect(TABLE_MAX_COLS).toBe(20);
    expect(overTableCap(60, 20)).toBe(false);
    expect(overTableCap(61, 20)).toBe(true);
    expect(overTableCap(60, 21)).toBe(true);
  });

  // What one repaint round trip is allowed to carry: a picture weighs its
  // base64, a table the JSON of its cells.
  it("weighs a table by the JSON of its cells", () => {
    expect(payloadBytes(table)).toBe(JSON.stringify(table.cells).length);
  });
});

describe("sourceLabel", () => {
  it("names ranges, charts and tables", () => {
    expect(sourceLabel(src, "range")).toBe("Model!B4:F12");
    expect(sourceLabel({ ...src, ref: "Revenue bridge" }, "chart")).toBe(
      "Model: Revenue bridge",
    );
    expect(sourceLabel(src, "table")).toBe("Model!B4:F12 table");
  });
});
