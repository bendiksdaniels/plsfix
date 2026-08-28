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
  sourceLabel,
  type LinkTag,
  type Payload,
} from "./model";

const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "SMT_LINK_0123abcd",
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
    expect(anchorName("0123abcd".repeat(4))).toBe("SMT_LINK_0123abcd");
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
          anchor: "SMT_LINK_0123abcd",
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
  it("round-trips through UTF-8 bytes", () => {
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
    expect(() => decodePayload(new TextEncoder().encode('{"v":1}'))).toThrow(
      /payload/,
    );
  });
});

describe("sourceLabel", () => {
  it("names ranges and charts", () => {
    expect(sourceLabel(src, "range")).toBe("Model!B4:F12");
    expect(sourceLabel({ ...src, ref: "Revenue bridge" }, "chart")).toBe(
      "Model: Revenue bridge",
    );
  });
});
